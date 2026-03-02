const { chromium } = require('playwright');
const { readExcel, markExcel } = require('./excel-helper');
const { processAllPatients, loadExcelData } = require('./assessment-processor');
const path = require('path');
require('dotenv').config();

// Helper function to convert Excel serial date to MM/DD/YYYY
function excelDateToMMDDYYYY(excelDate) {
  // If already in MM/DD/YYYY format, return as is
  if (typeof excelDate === 'string' && excelDate.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
    return excelDate;
  }
  
  // Convert Excel serial number to JavaScript Date
  // Excel date serial number starts from 1/1/1900
  const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
  const jsDate = new Date(excelEpoch.getTime() + excelDate * 86400000);
  
  // Format as MM/DD/YYYY
  const month = String(jsDate.getMonth() + 1).padStart(2, '0');
  const day = String(jsDate.getDate()).padStart(2, '0');
  const year = jsDate.getFullYear();
  
  return `${month}/${day}/${year}`;
} 

// ===== Configuration =====
const CONFIG = {
  BASE_URL: 'https://my.qhslab.com',
  CLIENT_FILE: path.resolve(__dirname, 'clients', 'FMG 09.22.2025.xlsx'),
  DEFAULT_PROVIDER: 'REGINALD JEROME APRN',
  DEFAULT_INSURANCE: 'AvMed',
  HEADLESS: false,
  SLOW_MO: 2000,
  
  TIMEOUTS: {
    PAGE_LOAD: 30000,
    ELEMENT_WAIT: 10000,
    SHORT: 1000,
    MEDIUM: 2000,
    LONG: 3000
  },
  
  ASSESSMENT_TYPES: {
    HEALTH: 'Health Assessment',
    PHQ_GAD16: 'PHQ-GAD16 Health Assessment'
  },
  
  STATUS: {
    SENT: 'Sent',
    ERROR: 'need to add demo',
    ALREADY: 'Already',
    NEED_DEMO: 'failed to fetch',
    UNABLE: 'Unable',
    PATIENT_NOT_FOUND: 'Patient not found'
  },
  
  CREDS: {
    email: process.env.QHSLAB_EMAIL || 'adam.nelson@medviz.ai',
    password: process.env.QHSLAB_PASSWORD || 'medviz@741'
  }
};

// ===== Verification Functions =====
async function verifyLogin(page) {
  try {
    console.log('🔍 Verifying login...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Check if we're still on login page (login failed)
    const currentUrl = page.url();
    if (currentUrl.includes('/login')) {
      // Check if there's an error message
      const errorMessages = await page.locator('[role="alert"], .error, .MuiAlert-root').count();
      if (errorMessages > 0) {
        throw new Error('Login failed - error message detected');
      }
      // Wait a bit more to see if page redirects
      await page.waitForTimeout(2000);
      const newUrl = page.url();
      if (newUrl.includes('/login')) {
        throw new Error('Login verification failed - still on login page');
      }
    }
    
    // Check if page is closed
    if (page.isClosed()) {
      throw new Error('Page was closed during login verification');
    }
    
    console.log('✅ Login verified successfully');
    return true;
  } catch (error) {
    console.error('❌ Login verification failed:', error.message);
    throw error;
  }
}

async function verifySearchPage(page) {
  try {
    console.log('🔍 Verifying search page...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during search page verification');
    }
    
    // Check if we're on the accounts/search page
    const currentUrl = page.url();
    if (!currentUrl.includes('/accounts')) {
      throw new Error(`Search page verification failed - current URL: ${currentUrl}`);
    }
    
    // Check if the accounts table is visible
    try {
      const customIdCell = page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' });
      await customIdCell.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      console.log('✅ Search page verified successfully');
      return true;
    } catch (e) {
      throw new Error('Search page verification failed - accounts table not found');
    }
  } catch (error) {
    console.error('❌ Search page verification failed:', error.message);
    throw error;
  }
}

async function verifyAccountSelected(page, accountName) {
  try {
    console.log(`🔍 Verifying account selection: ${accountName}...`);
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during account selection verification');
    }
    
    // Check if we're no longer on the accounts list page (should be on account detail page)
    const currentUrl = page.url();
    if (currentUrl.includes('/accounts') && !currentUrl.includes('/accounts/')) {
      // Still on list page - account selection may have failed
      // Check if account row is still visible (might mean click didn't work)
      const accountRows = await page.locator('tbody tr').count();
      if (accountRows > 0) {
        throw new Error('Account selection verification failed - still on accounts list page');
      }
    }
    
    console.log('✅ Account selection verified successfully');
    return true;
  } catch (error) {
    console.error('❌ Account selection verification failed:', error.message);
    throw error;
  }
}

async function verifyPatientSelected(page, dob) {
  try {
    console.log(`🔍 Verifying patient selection: ${dob}...`);
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during patient selection verification');
    }
    
    // Check if we're on a patient detail page (not on accounts list)
    const currentUrl = page.url();
    if (currentUrl.includes('/accounts') && !currentUrl.includes('/accounts/')) {
      throw new Error('Patient selection verification failed - still on accounts page');
    }
    
    // Check if patient content area is visible
    try {
      const contentArea = page.locator('#contentArea');
      await contentArea.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      console.log('✅ Patient selection verified successfully');
      return true;
    } catch (e) {
      throw new Error('Patient selection verification failed - patient content area not found');
    }
  } catch (error) {
    console.error('❌ Patient selection verification failed:', error.message);
    throw error;
  }
}

async function verifyAssessmentButtonClicked(page) {
  try {
    console.log('🔍 Verifying assessment button click...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during assessment button verification');
    }
    
    // Check if "Create Assessment" menu item is visible
    try {
      const createAssessmentMenu = page.getByRole('menuitem', { name: 'Create Assessment' });
      await createAssessmentMenu.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      console.log('✅ Assessment button click verified successfully');
      return true;
    } catch (e) {
      throw new Error('Assessment button verification failed - Create Assessment menu not found');
    }
  } catch (error) {
    console.error('❌ Assessment button verification failed:', error.message);
    throw error;
  }
}

async function verifyAssessmentTypeSelected(page, assessmentType) {
  try {
    console.log(`🔍 Verifying assessment type selection: ${assessmentType}...`);
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during assessment type verification');
    }
    
    // Check if appointment form fields are visible (indicates assessment type was selected)
    try {
      const appointmentProvider = page.getByLabel('Appointment Provider');
      await appointmentProvider.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      console.log('✅ Assessment type selection verified successfully');
      return true;
    } catch (e) {
      throw new Error('Assessment type verification failed - appointment form not found');
    }
  } catch (error) {
    console.error('❌ Assessment type verification failed:', error.message);
    throw error;
  }
}

async function verifyFormFilled(page, patientData) {
  try {
    console.log('🔍 Verifying form fill...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during form fill verification');
    }
    
    // Check if required fields have values
    const dateFields = page.locator('input[placeholder*="MM/DD/YYYY"]');
    const dateCount = await dateFields.count();
    
    if (dateCount === 0) {
      throw new Error('Form fill verification failed - no date fields found');
    }
    
    // Check if at least one date field has a value
    let hasDateValue = false;
    for (let i = 0; i < dateCount; i++) {
      const value = await dateFields.nth(i).inputValue();
      if (value && value.trim() !== '') {
        hasDateValue = true;
        break;
      }
    }
    
    if (!hasDateValue) {
      throw new Error('Form fill verification failed - no date values found');
    }
    
    console.log('✅ Form fill verified successfully');
    return true;
  } catch (error) {
    console.error('❌ Form fill verification failed:', error.message);
    throw error;
  }
}

async function verifyAssessmentSent(page) {
  try {
    console.log('🔍 Verifying assessment sent...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    if (page.isClosed()) {
      throw new Error('Page was closed during send verification');
    }
    
    // Check if Send button is no longer visible (form was submitted)
    try {
      const sendButton = page.locator('button').filter({ hasText: /^Send$/ });
      const isVisible = await sendButton.isVisible({ timeout: 2000 });
      if (isVisible) {
        throw new Error('Send verification failed - Send button still visible');
      }
      console.log('✅ Assessment send verified successfully');
      return true;
    } catch (e) {
      // If button not found, it might mean form was submitted successfully
      if (e.message.includes('still visible')) {
        throw e;
      }
      console.log('✅ Assessment send verified successfully (Send button no longer visible)');
      return true;
    }
  } catch (error) {
    console.error('❌ Assessment send verification failed:', error.message);
    throw error;
  }
}

// ===== Core Functions =====
async function login(page) {
  try {
    console.log('🌐 Navigating to login page...');
    await page.goto(`${CONFIG.BASE_URL}/login`, { waitUntil: 'domcontentloaded', timeout: CONFIG.TIMEOUTS.PAGE_LOAD });
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed during login navigation');
    }
    
    const flexDiv = page.locator('div.MuiGrid-root.MuiGrid-container.MuiGrid-align-items-xs-center.MuiGrid-justify-content-xs-center').first();
    await flexDiv.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
    await flexDiv.click();
    
    await flexDiv.locator('input').nth(0).fill(CONFIG.CREDS.email);
    await flexDiv.locator('input').nth(1).fill(CONFIG.CREDS.password);
    await page.getByRole('button', { name: 'Login' }).click();
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Check if page is still valid after login
    if (page.isClosed()) {
      throw new Error('Page was closed after login');
    }
    
    console.log('✅ Login completed');
    
    // Verify login completed successfully
    await verifyLogin(page);
  } catch (error) {
    console.error('❌ Error during login:', error.message);
    throw error;
  }
}

async function openSearch(page) {
  try {
    console.log('🔍 Navigating to patients search page...');
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed before navigation');
    }
    
    await page.goto(`${CONFIG.BASE_URL}/6oQ5FvCBDUC5CiIrutgARg/accounts`, { waitUntil: 'domcontentloaded', timeout: CONFIG.TIMEOUTS.PAGE_LOAD });
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    // Check if page is still valid after navigation
    if (page.isClosed()) {
      throw new Error('Page was closed after navigation');
    }
    
    // Check for "OPEN ROOT PAGE" button and click it if present (Page not found error)
    try {
      const openRootPageButton = page.locator('button:has-text("OPEN ROOT PAGE")').first();
      const isVisible = await openRootPageButton.isVisible({ timeout: 3000 });
      if (isVisible) {
        console.log('🔵 Clicking "OPEN ROOT PAGE" button...');
        await openRootPageButton.click();
        await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
        console.log('✅ "OPEN ROOT PAGE" button clicked successfully');
        
        // Wait a bit more for the page to load after clicking
        await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
      }
    } catch (e) {
      console.log('ℹ️ No "OPEN ROOT PAGE" button found');
    }
    
    // Check for "Go to page" button and click it if present
    try {
      const goToPageButton = page.locator('button:has-text("Go to page")').first();
      const isVisible = await goToPageButton.isVisible({ timeout: 3000 });
      if (isVisible) {
        console.log('🔵 Clicking "Go to page" button...');
        await goToPageButton.click();
        await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
        console.log('✅ "Go to page" button clicked successfully');
      }
    } catch (e) {
      console.log('ℹ️ No "Go to page" button found or already clicked');
    }
    
    console.log('✅ Search page opened successfully');
    
    // Verify search page opened successfully
    await verifySearchPage(page);
  } catch (error) {
    console.error('❌ Error navigating to search page:', error.message);
    throw error;
  }
}

async function selectAccountOnce(page, accountName, customId) {
  try {
    console.log(`🔍 Filtering by Custom ID for account: ${accountName}`);
    
    if (!customId) {
      throw new Error(`No Custom ID provided for account: ${accountName}`);
    }
    
    console.log(`🔍 Using Custom ID: ${customId} for account: ${accountName}`);
    
    // Check if page is still valid before starting
    if (page.isClosed()) {
      throw new Error('Page was closed before account selection');
    }
    
    // Wait a bit to ensure page is stable
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Click on Custom ID filter
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByPlaceholder('filter').click();
    
    // Fill in the Custom ID
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByPlaceholder('filter').fill(customId);
    
    // Check if page is still valid before proceeding
    if (page.isClosed()) {
      throw new Error('Page was closed before account selection');
    }
    
    // Wait for the filtered results to load
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    // Try to find and click on the account row (single click instead of double click)
    try {
      // Look for any cell that contains the account name or similar text
      const accountCell = page.locator('td').filter({ hasText: new RegExp(accountName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') }).first();
      await accountCell.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      await accountCell.click();
      console.log(`✅ Account clicked successfully using Custom ID ${customId}: ${accountName}`);
    } catch (error) {
      // Fallback: try to click on the first row if specific account not found
      console.log('⚠️ Specific account not found, trying first available row...');
      const firstRow = page.locator('tbody tr').first();
      await firstRow.click();
      console.log(`✅ First available account clicked using Custom ID ${customId}`);
    }
    
    // Wait a bit after clicking
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Check if page is still valid after account selection
    if (page.isClosed()) {
      throw new Error('Page was closed after account selection');
    }
    
    console.log('✅ Account selection completed successfully');
    
    // Verify account selection completed successfully
    await verifyAccountSelected(page, accountName);
  } catch (error) {
    console.log(`❌ Error applying account filter for ${accountName}:`, error.message);
    throw error;
  }
}

async function selectPatientByDOB(page, dob) {
  const dobString = String(dob);
  
  try {
    // Click on ISP cell to select the patient
    await page.getByRole('cell', { name: 'ISP' }).locator('div').first().click();

    // Filter by Patient Date of Birth
    await page.getByRole('cell', { name: 'Date of Birth Sort by Date of' }).getByRole('textbox').click();
    
    // Fill DOB
    await page.getByRole('textbox', { name: 'MM/DD/YYYY' }).fill(dobString);
    console.log(`✅ DOB filter filled with: ${dobString}`);

    // Select the patient
    await page.getByRole('cell', { name: 'ISP' }).locator('div').first().dblclick();

  } catch (error) {
    console.log('❌ Error selecting patient:', error.message);
    throw error;
  }
}

async function clickAssessmentPlusButton(page) {
  console.log('🔍 Looking for assessment plus button...');
  
  const allButtons = page.locator('.MuiButtonBase-root.MuiIconButton-root');
  const buttonCount = await allButtons.count();
  
  for (let i = 0; i < buttonCount; i++) {
    try {
      const button = allButtons.nth(i);
      if (!await button.isVisible()) continue;
      
      const ariaLabel = await button.getAttribute('aria-label');
      const parentText = await button.locator('..').textContent();
      
      if (ariaLabel?.toLowerCase().includes('add') || 
          ariaLabel?.toLowerCase().includes('assessment') ||
          parentText?.toLowerCase().includes('assessment') ||
          parentText?.toLowerCase().includes('add')) {
        await button.click();
        console.log(`✅ Assessment button clicked (button ${i})`);
        return;
      }
    } catch (e) {
      continue;
    }
  }
  
  throw new Error('Could not find assessment plus button');
}

async function clickCreateAssessment(page) {
  await page.getByRole('menuitem', { name: 'Create Assessment' }).click();
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Select assessment by button position: 1st = Health Assessment, 2nd = GAD 16
async function clickSelectForAssessment(page, title) {
  console.log(`🔍 Selecting assessment: "${title}"`);
 
  // Ensure the list is rendered (at least one "SELECT" visible)
  await page.getByRole('button', { name: /^select$/i })
            .first()
            .waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
 
  // Get all SELECT buttons
  const selectButtons = page.getByRole('button', { name: /^select$/i });
  const count = await selectButtons.count();
  
  if (!count) throw new Error('No SELECT buttons found.');
  
  console.log(`📊 Found ${count} SELECT buttons`);
  
  // Select based on assessment type
  if (title.toLowerCase().includes('health assessment') && !title.toLowerCase().includes('phq') && !title.toLowerCase().includes('gad')) {
    // Health Assessment = 1st button (index 0)
    if (count >= 1) {
      await selectButtons.nth(0).click();
      console.log(`✅ Selected Health Assessment (1st button)`);
      return;
    }
  } else if (title.toLowerCase().includes('phq') || title.toLowerCase().includes('gad')) {
    // PHQ-GAD16 = 2nd button (index 1)
    if (count >= 2) {
      await selectButtons.nth(1).click();
      console.log(`✅ Selected PHQ-GAD16 (2nd button)`);
      return;
    }
  }
  
  // Fallback: try to find by text if position-based selection fails
  console.log(`🔍 Fallback: Trying text-based selection`);
  const wanted = title.trim().toLowerCase();
  
  for (let i = 0; i < count; i++) {
    const btn = selectButtons.nth(i);
    const btnContainer = btn.locator('xpath=ancestor::div[contains(@class,"Mui")][1]');
    const containerText = (await btnContainer.innerText()).toLowerCase();
    
    console.log(`🔍 Button ${i} container text: "${containerText.substring(0, 100)}..."`);
    
    if (containerText.includes(wanted)) {
      await btn.click();
      console.log(`✅ Selected by text match: "${title}" (button ${i})`);
      return;
    }
  }
  
  // If we got here, log what we saw
  const seen = [];
  for (let i = 0; i < count; i++) {
    const btn = selectButtons.nth(i);
    const btnContainer = btn.locator('xpath=ancestor::div[contains(@class,"Mui")][1]');
    seen.push((await btnContainer.innerText()).split('\n')[0]);
  }
  throw new Error(`Could not find assessment "${title}". Available options: ${JSON.stringify(seen)}`);
}

async function selectAssessmentType(page, assessmentType) {
  console.log(`🔍 Selecting assessment type: ${assessmentType}`);
 
  try {
    await clickSelectForAssessment(page, assessmentType);
  } catch (err) {
    // Fallback (only if the title-based click fails)
    console.log('⚠️ Title-based selection failed, trying fallback by index…');
    const selectButtons = page.locator('button:has-text("Select")');
    const count = await selectButtons.count();
 
    if (assessmentType === CONFIG.ASSESSMENT_TYPES.HEALTH && count >= 1) {
      await selectButtons.first().click();
    } else if (assessmentType === CONFIG.ASSESSMENT_TYPES.PHQ_GAD16 && count >= 2) {
      await selectButtons.nth(1).click();
    } else {
      throw err;
    }
  }
}

function getTomorrowDate() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  return `${(tomorrow.getMonth() + 1).toString().padStart(2, '0')}/${tomorrow.getDate().toString().padStart(2, '0')}/${tomorrow.getFullYear()}`;
}

function getCurrentDate() {
  const now = new Date();
  return `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
}

async function getLastCompletedDate(page, dob) {
  try {
    console.log('🔍 Checking for recent Health Assessment...');
    
    // Look for the assessments table in the content area
    const contentArea = page.locator('#contentArea');
    if (await contentArea.isVisible()) {
      console.log('📊 Found content area');
      
      // Look for the table near the 'Assessments' text
      const assessmentsSection = contentArea.getByText('Assessments');
      if (await assessmentsSection.isVisible()) {
        console.log('📊 Found Assessments section');
        
        // Find the table in this section - try multiple approaches
        let table = null;
        
        // Approach 1: Look for table after the Assessments text
        table = assessmentsSection.locator('..').locator('table').first();
        if (!(await table.isVisible())) {
          // Approach 2: Look for table in the same parent container
          table = assessmentsSection.locator('..').locator('..').locator('table').first();
        }
        if (!(await table.isVisible())) {
          // Approach 3: Look for any table in the content area
          table = contentArea.locator('table').first();
        }
        if (!(await table.isVisible())) {
          // Approach 4: Look for table with tbody
          table = contentArea.locator('table tbody').locator('..').first();
        }
        
        if (await table.isVisible()) {
          console.log('📊 Found assessments table');
          
          // Get all rows in the table
          const rows = await table.locator('tbody tr').all();
          let latest = null;
          
          console.log(`📊 Found ${rows.length} assessment rows to check`);
          
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cells = await row.locator('td').all();
            
            if (cells.length > 0) {
              console.log(`🔍 Row ${i + 1}: Checking all ${cells.length} columns for status...`);
              
              let statusText = '';
              
              // Check all columns for status information
              for (let j = 0; j < cells.length; j++) {
                const cell = cells[j];
                const cellText = (await cell.innerText()).trim().toLowerCase();
                const cellHTML = await cell.innerHTML();
                
                console.log(`🔍 Row ${i + 1}, Column ${j + 1}: "${cellText}"`);
                
                // Look for status keywords in any column
                if (cellText && (cellText.includes('completed') || cellText.includes('pending review') || cellText.includes('e-transfer') || cellText.includes('invited'))) {
                  statusText = cellText;
                  console.log(`✅ Row ${i + 1}: Found status "${statusText}" in column ${j + 1}`);
                  break;
                }
                
                // Also check for status elements within the cell
                const statusElements = await cell.locator('span, button, div, [class*="status"], [class*="tag"]').all();
                for (const element of statusElements) {
                  const text = (await element.innerText()).trim().toLowerCase();
                  if (text && (text.includes('completed') || text.includes('pending review') || text.includes('e-transfer') || text.includes('invited'))) {
                    statusText = text;
                    console.log(`✅ Row ${i + 1}: Found status element "${statusText}" in column ${j + 1}`);
                    break;
                  }
                }
                
                if (statusText) break;
              }
              
              console.log(`🔍 Row ${i + 1} final status: "${statusText}"`);
              
              // Check for any of the 3 status types: Completed, Pending Review, E-Transfer
              if (statusText.includes('completed') || statusText.includes('pending review') || statusText.includes('e-transfer')) {
                // Look for Order Date in the date columns
                let dateText = '';
                for (let j = 0; j < cells.length; j++) {
                  const cellText = (await cells[j].innerText()).trim();
                  // Look for date pattern MM/DD/YYYY
                  if (cellText.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
                    dateText = cellText;
                    break;
                  }
                }
                
                if (dateText) {
                  console.log(`📅 Found assessment (${statusText}) on: ${dateText}`);
                  try {
                    const [m, d, y] = dateText.split('/');
                    const dt = new Date(`${y}-${m}-${d}`);
                    if (!latest || dt > latest) latest = dt;
                  } catch (e) {
                    console.log(`⚠️ Error parsing date: ${dateText}`);
                  }
                }
              }
            }
          }
          
          console.log(`📅 Latest assessment (Completed/Pending Review/E-Transfer): ${latest ? latest.toDateString() : 'None found'}`);
          return latest;
        } else {
          console.log('📝 No table found in Assessments section');
        }
      } else {
        console.log('📝 No Assessments section found');
      }
    } else {
      console.log('📝 No content area found');
    }
    
    console.log('📝 No assessments table found - proceeding with Health Assessment');
    return null;
    
  } catch (error) {
    console.log(`⚠️ Error checking assessment history: ${error.message}`);
    return null;
  }
}

async function fillAppointmentForm(page, patientData) {
  console.log('📝 Filling appointment form...');
  console.log('🔍 DEBUG - patientData received:', JSON.stringify(patientData, null, 2));
  
  // Fill appointment date - use date from Excel if available, otherwise use default
  const appointmentDate = patientData?.['Appointment Date'] || getTomorrowDate();
  const dateSelectors = [
    '.MuiInputBase-root.jss744', 'input[placeholder="MM/DD/YYYY"]',
    'input[placeholder*="MM/DD/YYYY"]', 'input[type="text"]',
    'input[type="date"]', '.MuiInputBase-input'
  ];
  
  for (const selector of dateSelectors) {
    try {
      const element = page.locator(selector);
      if (await element.count() > 0 && await element.first().isVisible()) {
        await element.first().click();
        await page.waitForTimeout(500);
        await element.first().fill(appointmentDate);
        console.log(`✅ Appointment date filled: ${appointmentDate}${patientData?.['Appointment Date'] ? ' (from Excel)' : ' (default)'}`);
        break;
      }
    } catch (e) {
      continue;
    }
  }
  
  // Select provider with better error handling
  const provider = patientData?.['Scheduler'] || '';
  console.log(`Using provider: "${provider}"`);
  
  let providerFound = false;
  
  if (provider) {
    try {
      console.log(`🔍 Attempting to select provider: "${provider}"`);
      
      // Click provider dropdown
      await page.getByLabel('Appointment Provider').click();
      await page.waitForTimeout(1000);
      
      // Wait for dropdown to be visible
      try {
        await page.locator('[role="menuitem"]').first().waitFor({ state: 'visible', timeout: 5000 });
        console.log('✅ Provider dropdown opened successfully');
      } catch (e) {
        console.log('⚠️ Provider dropdown not visible, trying alternative approach...');
      }
      
      // Try multiple strategies for provider selection
      const providerTrimmed = provider.trim();
      
      // Strategy 1: Exact match
      try {
        const exactProvider = page.getByRole('menuitem', { name: providerTrimmed });
        if (await exactProvider.isVisible()) {
          await exactProvider.click();
          providerFound = true;
          console.log(`✅ Provider selected (exact match): ${providerTrimmed}`);
        }
      } catch (e) {
        console.log('🔍 Exact provider match not found, trying other strategies...');
      }
      
      // Strategy 2: With trailing space
      if (!providerFound) {
        try {
          const providerWithSpace = page.getByRole('menuitem', { name: providerTrimmed + ' ' });
          if (await providerWithSpace.isVisible()) {
            await providerWithSpace.click();
            providerFound = true;
            console.log(`✅ Provider selected (with space): ${providerTrimmed}`);
          }
        } catch (e) {
          console.log('🔍 Provider match with space not found...');
        }
      }
      
      // Strategy 3: Partial match
      if (!providerFound) {
        try {
          const allOptions = page.locator('[role="menuitem"]');
          const count = await allOptions.count();
          
          for (let i = 0; i < count; i++) {
            const option = allOptions.nth(i);
            const text = await option.textContent();
            
            if (text && text.trim().toLowerCase().includes(providerTrimmed.toLowerCase())) {
              await option.click();
              providerFound = true;
              console.log(`✅ Provider selected (partial match): "${text.trim()}"`);
              break;
            }
          }
        } catch (e) {
          console.log('🔍 Provider partial match failed:', e.message);
        }
      }
      
      if (!providerFound) {
        console.log(`⚠️ Provider option "${providerTrimmed}" not found after trying all strategies`);
        // Close dropdown if still open before random click - try multiple methods
        try {
          // Method 1: Press Escape
          await page.keyboard.press('Escape');
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        
        // Method 2: Click outside the dropdown to close it
        try {
          const providerField = page.getByLabel('Appointment Provider');
          if (await providerField.isVisible({ timeout: 1000 }).catch(() => false)) {
            // Click on the provider field itself to close dropdown
            await providerField.click({ force: true });
            await page.waitForTimeout(300);
          }
        } catch (e) {
          // Ignore errors
        }
        
        // Method 3: Click somewhere else on the page
        try {
          await page.mouse.click(100, 100);
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        
        // Wait a bit more to ensure dropdown is closed
        await page.waitForTimeout(500);
      }
      
    } catch (e) {
      console.log('⚠️ Error selecting provider:', e.message);
      console.log('⚠️ Could not select provider, skipping');
      // If provider data was provided but selection failed, close dropdown and perform random click
      if (provider && provider.trim() !== '') {
        // Close dropdown first
        try {
          await page.keyboard.press('Escape');
          await page.waitForTimeout(300);
        } catch (err) {
          // Ignore errors
        }
        try {
          await page.mouse.click(100, 100);
          await page.waitForTimeout(300);
        } catch (err) {
          // Ignore errors
        }
        await page.waitForTimeout(500);
      }
    }
  } else {
    console.log('⚠️ No provider data, skipping provider selection');
    // If no provider data provided, consider it as "found" (not an error case)
    providerFound = true;
  }
  
  // Select insurance with multiple strategies for large dropdowns
  let insurance = patientData?.['Primary Insurance Name'] || '';
  console.log(`🔍 DEBUG - Insurance data: "${insurance}"`);
  
  let optionFound = false;
  
  if (insurance && insurance.trim() !== '') {
    try {
      console.log(`🔍 Attempting to select insurance: "${insurance}"`);
      
      // Click insurance dropdown
      await page.getByLabel('Insurance').click();
      await page.waitForTimeout(1000);
      
      // Wait for dropdown to be visible with better timeout handling
      try {
        await page.locator('[role="menuitem"]').first().waitFor({ state: 'visible', timeout: 3000 });
        console.log('✅ Insurance dropdown opened successfully');
      } catch (e) {
        console.log('⚠️ Insurance dropdown not visible, trying alternative approach...');
        // Try to wait a bit more and check again
        await page.waitForTimeout(2000);
      }
      
      // Try multiple approaches to find the insurance option
      const insuranceTrimmed = insurance.trim();
      
      // Strategy 1: Try exact match
      try {
        const exactOption = page.getByRole('menuitem', { name: insuranceTrimmed });
        if (await exactOption.isVisible()) {
          await exactOption.click();
          optionFound = true;
          console.log(`✅ Insurance selected (exact match): ${insuranceTrimmed}`);
        }
      } catch (e) {
        console.log('🔍 Exact match not found, trying other strategies...');
      }
      
      // Strategy 2: Try with trailing space (common issue)
      if (!optionFound) {
        try {
          const optionWithSpace = page.getByRole('menuitem', { name: insuranceTrimmed + ' ' });
          if (await optionWithSpace.isVisible()) {
            await optionWithSpace.click();
            optionFound = true;
            console.log(`✅ Insurance selected (with space): ${insuranceTrimmed}`);
          }
        } catch (e) {
          console.log('🔍 Match with space not found...');
        }
      }
      
      // Strategy 3: Try exact text match first (more precise)
      if (!optionFound) {
        try {
          const allOptions = page.locator('[role="menuitem"]');
          const count = await allOptions.count();
          
          console.log(`🔍 Checking ${count} insurance options for exact match...`);
          
          for (let i = 0; i < count; i++) {
            const option = allOptions.nth(i);
            const text = await option.textContent();
            const textTrimmed = text ? text.trim() : '';
            
            console.log(`🔍 Option ${i + 1}: "${textTrimmed}"`);
            
            // Try exact match first
            if (textTrimmed === insuranceTrimmed) {
              try {
                await option.scrollIntoViewIfNeeded({ timeout: 1000 });
                await option.click();
                optionFound = true;
                console.log(`✅ Insurance selected (exact text match): "${textTrimmed}"`);
                break;
              } catch (scrollError) {
                console.log(`🔍 Scroll failed for exact match "${textTrimmed}", trying next...`);
                continue;
              }
            }
            
            // Try exact match with trailing space
            if (textTrimmed === insuranceTrimmed + ' ') {
              try {
                await option.scrollIntoViewIfNeeded({ timeout: 1000 });
                await option.click();
                optionFound = true;
                console.log(`✅ Insurance selected (exact with space): "${textTrimmed}"`);
                break;
              } catch (scrollError) {
                console.log(`🔍 Scroll failed for space match "${textTrimmed}", trying next...`);
                continue;
              }
            }
          }
        } catch (e) {
          console.log('🔍 Exact text match failed:', e.message);
        }
      }
      
      // Strategy 4: Try partial match only if exact match fails (more conservative)
      if (!optionFound) {
        try {
          const allOptions = page.locator('[role="menuitem"]');
          const count = await allOptions.count();
          
          console.log(`🔍 Trying partial match for: "${insuranceTrimmed}"`);
          
          for (let i = 0; i < count; i++) {
            const option = allOptions.nth(i);
            const text = await option.textContent();
            const textTrimmed = text ? text.trim() : '';
            
            // Only match if the text contains the full insurance name (not just part)
            if (textTrimmed && textTrimmed.toLowerCase().includes(insuranceTrimmed.toLowerCase())) {
              // Make sure it's not too short (avoid matching "BCBS" when looking for "BCBS SS COMM")
              if (textTrimmed.length >= insuranceTrimmed.length * 0.8) {
                try {
                  await option.scrollIntoViewIfNeeded({ timeout: 1000 });
                  await option.click();
                  optionFound = true;
                  console.log(`✅ Insurance selected (partial match): "${textTrimmed}"`);
                  break;
                } catch (scrollError) {
                  console.log(`🔍 Scroll failed for partial match "${textTrimmed}", trying next...`);
                  continue;
                }
              }
            }
          }
        } catch (e) {
          console.log('🔍 Partial match failed:', e.message);
        }
      }
      
      if (!optionFound) {
        console.log(`⚠️ Insurance option "${insuranceTrimmed}" not found after trying all strategies`);
        // Close dropdown if still open before random click - try multiple methods
        try {
          // Method 1: Press Escape
          await page.keyboard.press('Escape');
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        
        // Method 2: Click outside the dropdown to close it
        try {
          const insuranceField = page.getByLabel('Insurance');
          if (await insuranceField.isVisible({ timeout: 1000 }).catch(() => false)) {
            // Click on the insurance field itself to close dropdown
            await insuranceField.click({ force: true });
            await page.waitForTimeout(300);
          }
        } catch (e) {
          // Ignore errors
        }
        
        // Method 3: Click somewhere else on the page
        try {
          await page.mouse.click(100, 100);
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        
        // Wait a bit more to ensure dropdown is closed
        await page.waitForTimeout(500);
      }
      
    } catch (e) {
      console.log('⚠️ Error selecting insurance:', e.message);
      console.log('⚠️ Could not select insurance, skipping');
      // If insurance data was provided but selection failed, close dropdown and perform random click
      if (insurance && insurance.trim() !== '') {
        // Close dropdown first
        try {
          await page.keyboard.press('Escape');
          await page.waitForTimeout(300);
        } catch (err) {
          // Ignore errors
        }
        try {
          await page.mouse.click(100, 100);
          await page.waitForTimeout(300);
        } catch (err) {
          // Ignore errors
        }
        await page.waitForTimeout(500);
      }
    }
  } else {
    console.log('⚠️ No insurance data, skipping insurance selection');
    // If no insurance data provided, consider it as "found" (not an error case)
    optionFound = true;
  }
  
  // Schedule for Later functionality - COMMENTED OUT: Now sending directly
  // await page.getByLabel('Schedule for Later').check();
  // await page.getByPlaceholder('MM/DD/YYYY').nth(1).click();
  // await page.getByPlaceholder('MM/DD/YYYY').nth(1).fill(getCurrentDate());
  
  console.log('✅ Appointment form filled');
}

async function verifyAppointmentInHistory(page, patientData) {
  console.log('🔍 Verifying appointment in appointment history...');
  
  try {
    // Wait a moment for the appointment to be processed
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    // Use the same approach as the 30-day gap check - look for the assessments table
    const contentArea = page.locator('#contentArea');
    if (await contentArea.isVisible()) {
      console.log('📊 Found content area');
      
      // Look for the table near the 'Assessments' text (same as 30-day gap check)
      const assessmentsSection = contentArea.getByText('Assessments');
      if (await assessmentsSection.isVisible()) {
        console.log('📊 Found Assessments section');
        
        // Find the table in this section - try multiple approaches (same as 30-day gap check)
        let table = null;
        
        // Approach 1: Look for table after the Assessments text
        table = assessmentsSection.locator('..').locator('table').first();
        if (!(await table.isVisible())) {
          // Approach 2: Look for table in the same parent container
          table = assessmentsSection.locator('..').locator('..').locator('table').first();
        }
        if (!(await table.isVisible())) {
          // Approach 3: Look for any table in the content area
          table = contentArea.locator('table').first();
        }
        if (!(await table.isVisible())) {
          // Approach 4: Look for table with tbody
          table = contentArea.locator('table tbody').locator('..').first();
        }
        
        if (await table.isVisible()) {
          console.log('📊 Found assessments table for verification');
          
          // Get all rows in the table
          const rows = await table.locator('tbody tr').all();
          console.log(`📊 Found ${rows.length} assessment rows to check for new appointment`);
          
          // Look for the appointment by date patterns
          // Use appointment date from Excel if available, otherwise use default
          const appointmentDate = patientData?.['Appointment Date'] || getTomorrowDate();
          const currentDate = getCurrentDate();
          
          // Try to find appointment by date (appointment date from Excel or current date)
          const datePatterns = [
            appointmentDate,
            currentDate,
            appointmentDate.split('/')[1] + '/' + appointmentDate.split('/')[0] + '/' + appointmentDate.split('/')[2], // DD/MM/YYYY format
            currentDate.split('/')[1] + '/' + currentDate.split('/')[0] + '/' + currentDate.split('/')[2]  // DD/MM/YYYY format
          ];
          
          console.log(`🔍 Looking for appointment with dates: ${datePatterns.join(', ')}`);
          
          // Check each row for the new appointment
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cells = await row.locator('td').all();
            
            if (cells.length > 0) {
              // Check all columns for date patterns
              for (let j = 0; j < cells.length; j++) {
                const cellText = (await cells[j].innerText()).trim();
                
                // Look for date pattern MM/DD/YYYY or DD/MM/YYYY
                if (cellText.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
                  console.log(`🔍 Row ${i + 1}, Column ${j + 1}: Found date "${cellText}"`);
                  
                  // Check if this date matches any of our target dates
                  for (const datePattern of datePatterns) {
                    if (cellText === datePattern) {
                      console.log(`✅ Found new appointment with date: ${cellText}`);
                      return true;
                    }
                  }
                }
              }
            }
          }
          
          // Alternative: Look for any recent appointment (within next few days)
          console.log('🔍 Checking for any recent appointments...');
          const recentDates = [];
          for (let i = 0; i < 7; i++) {
            const date = new Date();
            date.setDate(date.getDate() + i);
            const dateStr = `${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}/${date.getFullYear()}`;
            recentDates.push(dateStr);
          }
          
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cells = await row.locator('td').all();
            
            if (cells.length > 0) {
              for (let j = 0; j < cells.length; j++) {
                const cellText = (await cells[j].innerText()).trim();
                
                if (cellText.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
                  for (const recentDate of recentDates) {
                    if (cellText === recentDate) {
                      console.log(`✅ Found recent appointment with date: ${cellText}`);
                      return true;
                    }
                  }
                }
              }
            }
          }
          
          console.log('❌ No new appointment found in assessments table');
          return false;
        } else {
          console.log('📝 No assessments table found for verification');
        }
      } else {
        console.log('📝 No Assessments section found for verification');
      }
    } else {
      console.log('📝 No content area found for verification');
    }
    
    console.log('❌ No appointment found in appointment history');
    return false;
    
  } catch (error) {
    console.log('⚠️ Error verifying appointment in history:', error.message);
    return false;
  }
}

async function createAssessment(page, assessmentType, patientData, accountName) {
  console.log(`🔧 Creating ${assessmentType}...`);
  
  try {
    await clickAssessmentPlusButton(page);
    
    // Verify assessment button clicked successfully
    await verifyAssessmentButtonClicked(page);
    
    await clickCreateAssessment(page);
    await selectAssessmentType(page, assessmentType);
    
    // Verify assessment type selected successfully
    await verifyAssessmentTypeSelected(page, assessmentType);
    
    await fillAppointmentForm(page, patientData);
    
    // Verify form filled successfully
    await verifyFormFilled(page, patientData);
    
    console.log('📤 Sending assessment...');
    await page.locator('button').filter({ hasText: /^Send$/ }).click();
    console.log('✅ Assessment sent');
    
    // Verify assessment sent successfully
    await verifyAssessmentSent(page);
    
    // Add random click on left side of screen to maintain session activity
    try {
      console.log('🖱️ Adding random click to maintain session...');
      await page.click('body', { position: { x: 100, y: 200 } });
      await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
      console.log('✅ Random click completed');
    } catch (error) {
      console.log('⚠️ Random click failed, continuing anyway:', error.message);
    }
    
    // Verify appointment in appointment history
    console.log('🔍 Verifying appointment in appointment history...');
    const appointmentFound = await verifyAppointmentInHistory(page, patientData);
    
    // Determine final status based on verification
    const currentDate = getCurrentDate();
    let finalStatus;
    
    if (appointmentFound) {
      finalStatus = CONFIG.STATUS.SENT;
      console.log('✅ Appointment verified in history - Status: Sent');
    } else {
      finalStatus = CONFIG.STATUS.UNABLE;
      console.log('❌ Appointment not found in history - Status: Unable');
    }
    
    const result = { result: finalStatus, lastOrderDate: currentDate };
    
    return result;
    
  } catch (error) {
    console.log(`❌ Error creating assessment: ${error.message}`);
    throw error;
  }
}

async function processPatient(page, dob, apptStr, provider, insurance, assessmentType, isHealthAssessment, dateStr, accountName, customId, appointmentDate = null) {
  try {
    console.log(`\n👤 Processing patient DOB: ${dob} (Account: ${accountName})`);
    
    if (!dob || dob === '') {
      console.log(`⚠️ Patient has no DOB - need to add demo`);
      return CONFIG.STATUS.NEED_DEMO;
    }
    
    // Note: dob variable contains DOB value from MRN column in Excel
    
    // Select account first
    await selectAccountOnce(page, accountName, customId);
    
    // Search for DOB immediately after account selection
    console.log(`🔍 Searching for DOB: ${dob}`);
    try {
      // Click on DOB filter field using the specific locator
      await page.getByRole('cell', { name: 'Date of Birth Sort by Date of' }).getByRole('textbox').click();
      
      // Fill in the DOB
      await page.getByRole('textbox', { name: 'MM/DD/YYYY' }).fill(dob);
      
      console.log(`✅ DOB filter filled with: ${dob}`);
      
      // Click Apply button to trigger search
      await page.getByRole('button', { name: 'Apply' }).click();
      console.log('✅ Clicked Apply button to search');
      
      // Wait for results to load
      await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
      
      // Check if patient exists by looking for ISP cell
      const ispCell = page.getByRole('cell', { name: 'ISP' });
      const ispCount = await ispCell.count();
      
      if (ispCount === 0) {
        console.log(`❌ Patient DOB ${dob} not found in system`);
        return CONFIG.STATUS.NEED_DEMO;
      }
      
      // If multiple patients found, select the first one
      if (ispCount > 1) {
        console.log(`⚠️ Multiple patients found for DOB ${dob}, selecting the first one (${ispCount} results)`);
      }
      
      // Double-click on the first ISP cell to select patient
      await ispCell.first().dblclick();
      console.log(`✅ Patient selected by double-clicking ISP cell: ${dob} (${ispCount} result${ispCount > 1 ? 's' : ''} found)`);
      
      // Wait for patient dashboard to load
      await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
      
      // Verify patient selection completed successfully
      await verifyPatientSelected(page, dob);
      
    } catch (error) {
      console.log(`❌ Error selecting patient by DOB: ${error.message}`);
      
      // Check if it's a "patient not found" error
      if (error.message.includes('not found') || error.message.includes('DOB')) {
        console.log(`⚠️ Patient DOB ${dob} not found in system - need to add demo`);
        return CONFIG.STATUS.NEED_DEMO;
      }
      
      throw error;
    }
    
    // Check for recent assessments (Health Assessment only)
    if (isHealthAssessment) {
      console.log('🔍 Checking for recent Health Assessment...');
      try {
        const lastCompletedDate = await getLastCompletedDate(page, dob);
        if (lastCompletedDate) {
          const thirtyDaysAgo = new Date();
          thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
          
          console.log(`📅 Last completed assessment: ${lastCompletedDate.toDateString()}`);
          console.log(`📅 30 days ago: ${thirtyDaysAgo.toDateString()}`);
          
          if (lastCompletedDate > thirtyDaysAgo) {
            console.log('✅ Recent assessment found within 30 days - skipping Health Assessment');
            return CONFIG.STATUS.ALREADY;
          } else {
            console.log('⏰ No recent assessment found - proceeding with Health Assessment');
          }
        } else {
          console.log('📝 No completed assessments found - proceeding with Health Assessment');
        }
      } catch (error) {
        console.log('⚠️ Error checking recent assessments:', error.message);
        console.log('❌ Cannot determine assessment status - SKIPPING to prevent duplicates');
        return CONFIG.STATUS.UNABLE;
      }
    }
    
    // Create assessment
    const patientData = {
      'Scheduler': provider,
      'Primary Insurance Name': insurance, // Use standard key for form
      'Last Assessment Date': isHealthAssessment ? null : undefined,
      'Appointment Date': appointmentDate
    };
    
    const assessmentResult = await createAssessment(page, assessmentType, patientData, accountName);
    return assessmentResult.result;
    
  } catch (error) {
    console.error(`❌ Error processing patient ${dob}:`, error.message);
    
    // Check if it's a "patient not found" error
    if (error.message.includes('not found') || error.message.includes('DOB')) {
      console.log(`⚠️ Patient DOB ${dob} not found in system - need to add demo`);
      return CONFIG.STATUS.NEED_DEMO;
    }
    
    await page.screenshot({ path: `error-${dob}.png` }).catch(() => {});
    return CONFIG.STATUS.ERROR;
  }
}

function getBrowserArgs() {
  return [
    '--no-sandbox',
    '--disable-setuid-sandbox',
    '--disable-dev-shm-usage',
    '--disable-web-security',
    '--disable-extensions',
    '--no-first-run',
    '--disable-default-apps',
    '--disable-sync',
    '--disable-gpu',
    '--disable-translate',
    '--hide-scrollbars',
    '--mute-audio',
    '--disable-blink-features=AutomationControlled',
    '--disable-features=VizDisplayCompositor',
    '--disable-background-timer-throttling',
    '--disable-backgrounding-occluded-windows',
    '--disable-renderer-backgrounding',
    '--disable-ipc-flooding-protection',
    '--disable-popup-blocking',
    '--disable-prompt-on-repost',
    '--disable-web-resources',
    '--enable-automation',
    '--password-store=basic',
    '--use-mock-keychain',
    '--disable-logging',
    '--disable-dev-tools',
    '--disable-extensions-file-access-check',
    '--disable-extensions-http-throttling',
    '--aggressive-cache-discard',
    '--memory-pressure-off',
    '--max_old_space_size=4096'
  ];
}

function setupPageListeners(page) {
  page.on('crash', () => console.log('❌ Page crashed'));
  page.on('close', () => console.log('❌ Page closed unexpectedly'));
  page.on('error', (error) => console.log('❌ Page error:', error.message));
  page.on('console', (msg) => {
    if (msg.type() === 'error') console.log('❌ Browser console error:', msg.text());
  });
}


async function navigateBackToAccounts(page) {
  try {
    console.log('🔄 Navigating back to accounts page...');
    
    // Click on Account@3x Accounts button
    await page.getByRole('button', { name: 'Account@3x Accounts' }).click();
    console.log('✅ Clicked Account@3x Accounts button');
    
    // Wait a moment for the page to load
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Click on Custom ID sort button (4th button)
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByRole('button').nth(3).click();
    console.log('✅ Clicked Custom ID sort button');
    
    // Wait a moment
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Click on Custom ID filter placeholder
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByPlaceholder('filter').click();
    console.log('✅ Clicked Custom ID filter placeholder');
    
    console.log('✅ Successfully navigated back to accounts page');
  } catch (error) {
    console.error('❌ Error navigating back to accounts:', error.message);
    throw error;
  }
}


// ===== Main Execution =====
(async () => {
  try {
    console.log('🚀 Starting assessment automation script...');
    console.log(`📁 Excel file path: ${CONFIG.CLIENT_FILE}`);
    
    const fs = require('fs');
    if (!fs.existsSync(CONFIG.CLIENT_FILE)) {
      console.error(`❌ Excel file not found at: ${CONFIG.CLIENT_FILE}`);
      process.exit(1);
    }
    console.log(`✅ Excel file exists: ${CONFIG.CLIENT_FILE}`);
    
    const data = await loadExcelData(CONFIG.CLIENT_FILE);
    console.log(`📊 Loaded ${data.gad16.length} GAD16 patients, ${data.health.length} HealthAssessment patients, and ${data.appointments.length} appointments`);

    console.log(`\n🎯 PROCESSING ALL PATIENTS: Sending assessments to all patients...`);
    console.log(`📝 Status updates will be saved to Excel file after each patient\n`);
    
    const result = await processAllPatients(CONFIG.CLIENT_FILE);
    
    console.log('\n🎉 All patients processed successfully!');
    console.log(`📊 Processing summary:`, result);

  } catch (error) {
    console.error('\n❌ Fatal error:', error.message);
    console.error('Stack trace:', error.stack);
    process.exit(1);
  }
})();