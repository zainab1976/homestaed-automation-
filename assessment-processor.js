const { chromium } = require('playwright');
const { readExcel, markExcel } = require('./excel-helper');
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
  DEFAULT_PROVIDER: 'REGINALD JEROME APRN',
  DEFAULT_INSURANCE: 'AvMed',
  HEADLESS: false,
  SLOW_MO: 2000,
  
  TIMEOUTS: {
  PAGE_LOAD: 30000,
    ELEMENT_WAIT: 5000,
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
    ERROR: 'Need to add demo',
    ALREADY: 'Already',
    NEED_DEMO: 'failed to fetch ',
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

// Advanced function to select from large dropdowns with multiple strategies
async function selectFromLargeDropdown(page, dropdownTrigger, optionName, options = {}) {
  const {
    maxScrollAttempts = 15,
    scrollStep = 200,
    waitTime = 300,
    searchStrategy = 'scroll' // 'scroll', 'type', 'fuzzy'
  } = options;

  try {
    console.log(`🔍 Selecting "${optionName}" from large dropdown...`);
    
    // Step 1: Open dropdown
    await dropdownTrigger.click();
    await page.waitForTimeout(500);
    
    // Step 2: Try different strategies based on dropdown type
    let optionFound = false;
    
    // Strategy 1: Direct search (if dropdown supports typing)
    if (searchStrategy === 'type' || searchStrategy === 'both') {
      try {
        const searchInput = page.locator('input[type="text"], input[placeholder*="search"], input[placeholder*="filter"]').first();
        if (await searchInput.isVisible()) {
          await searchInput.fill(optionName);
          await page.waitForTimeout(500);
          
          const option = page.getByRole('menuitem', { name: optionName });
          if (await option.isVisible()) {
            await option.click();
            optionFound = true;
            console.log(`✅ Selected "${optionName}" using search input`);
          }
        }
      } catch (e) {
        console.log('🔍 Search input not found, trying scroll strategy...');
      }
    }
    
    // Strategy 2: Scroll and find
    if (!optionFound && (searchStrategy === 'scroll' || searchStrategy === 'both')) {
      let scrollAttempt = 0;
      
      while (!optionFound && scrollAttempt < maxScrollAttempts) {
        try {
          // Try to find the option
          const option = page.getByRole('menuitem', { name: optionName });
          const isVisible = await option.isVisible();
          
          if (isVisible) {
            await option.scrollIntoViewIfNeeded();
            await option.click();
            optionFound = true;
            console.log(`✅ Selected "${optionName}" after ${scrollAttempt} scroll attempts`);
            break;
          }
        } catch (e) {
          // Option not found, continue scrolling
        }
        
        // Scroll down to load more options
        try {
          const dropdown = page.locator('[role="listbox"], .MuiMenu-list, .MuiSelect-select, .MuiAutocomplete-listbox').first();
          await dropdown.evaluate(el => el.scrollTop += scrollStep);
          await page.waitForTimeout(waitTime);
          scrollAttempt++;
        } catch (e) {
          break;
        }
      }
    }
    
    // Strategy 3: Fuzzy search (partial match)
    if (!optionFound && searchStrategy === 'fuzzy') {
      try {
        const allOptions = page.locator('[role="menuitem"], .MuiMenuItem-root, .MuiAutocomplete-option');
        const count = await allOptions.count();
        
        for (let i = 0; i < count; i++) {
          const option = allOptions.nth(i);
          const text = await option.textContent();
          
          if (text && text.toLowerCase().includes(optionName.toLowerCase())) {
            await option.scrollIntoViewIfNeeded();
            await option.click();
            optionFound = true;
            console.log(`✅ Selected "${optionName}" using fuzzy match: "${text}"`);
            break;
          }
        }
      } catch (e) {
        console.log('🔍 Fuzzy search failed:', e.message);
      }
    }
    
    if (!optionFound) {
      throw new Error(`Option "${optionName}" not found after trying all strategies`);
    }
    
  } catch (error) {
    console.log(`❌ Error selecting from dropdown: ${error.message}`);
    throw error;
  }
}

async function login(page) {
  try {
    console.log('🌐 Navigating to login page...');
    await page.goto(`${CONFIG.BASE_URL}/login`, { waitUntil: 'domcontentloaded', timeout: CONFIG.TIMEOUTS.PAGE_LOAD });
    
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
    
    if (page.isClosed()) {
      throw new Error('Page was closed before navigation');
    }
    
    await page.goto(`${CONFIG.BASE_URL}/6oQ5FvCBDUC5CiIrutgARg/accounts`, { waitUntil: 'domcontentloaded', timeout: CONFIG.TIMEOUTS.PAGE_LOAD });
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    if (page.isClosed()) {
      throw new Error('Page was closed after navigation');
    }
    
    // Check for "OPEN ROOT PAGE" button and click it if present
    try {
      const openRootPageButton = page.locator('button:has-text("OPEN ROOT PAGE")').first();
      const isVisible = await openRootPageButton.isVisible({ timeout: 3000 });
      if (isVisible) {
        console.log('🔵 Clicking "OPEN ROOT PAGE" button...');
        await openRootPageButton.click();
        await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
        console.log('✅ "OPEN ROOT PAGE" button clicked successfully');
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
      console.log(`⚠️ No Custom ID provided for account: ${accountName}`);
      console.log(`🔍 Attempting to find account by name only...`);
      
      // Try to find account by name without Custom ID filter
      try {
        // Look for account name in the table
        const accountCell = page.locator('td').filter({ hasText: new RegExp(accountName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') }).first();
        await accountCell.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
        await accountCell.click();
        console.log(`✅ Account clicked successfully by name: ${accountName}`);
        
        await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
        console.log('✅ Account selection completed successfully (by name)');
        return;
      } catch (nameError) {
        console.log(`❌ Could not find account by name: ${nameError.message}`);
        throw new Error(`No Custom ID provided for account: ${accountName} and account not found by name`);
      }
    }
    
    console.log(`🔍 Using Custom ID: ${customId} for account: ${accountName}`);
    
    if (page.isClosed()) {
      throw new Error('Page was closed before account selection');
    }
    
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
    // Click on Custom ID filter
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByPlaceholder('filter').click();
    
    // Fill in the Custom ID
    await page.getByRole('cell', { name: 'Custom ID Sort by Custom ID' }).getByPlaceholder('filter').fill(customId);
    
    if (page.isClosed()) {
      throw new Error('Page was closed before account selection');
    }
    
    await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
    
    // Try to find and click on the account row
    try {
      const accountCell = page.locator('td').filter({ hasText: new RegExp(accountName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') }).first();
      await accountCell.waitFor({ state: 'visible', timeout: CONFIG.TIMEOUTS.ELEMENT_WAIT });
      await accountCell.click();
      console.log(`✅ Account clicked successfully using Custom ID ${customId}: ${accountName}`);
    } catch (error) {
      console.log('⚠️ Specific account not found, trying first available row...');
      const firstRow = page.locator('tbody tr').first();
      await firstRow.click();
      console.log(`✅ First available account clicked using Custom ID ${customId}`);
    }
    
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);
    
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

    // Click Apply button to trigger search
    await page.getByRole('button', { name: 'Apply' }).click();
    console.log('✅ Clicked Apply button to search');
    
    // Wait for results to load
    await page.waitForTimeout(CONFIG.TIMEOUTS.SHORT);

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

async function clickSelectForAssessment(page, title) {
  console.log(`🔍 Selecting assessment: "${title}"`);
 
  // Ensure the list is rendered
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

// Helper function to perform a random click within the assessment form box
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
  
  // Select provider
  const provider = patientData?.['Scheduler'] || '';
  console.log(`Using provider: "${provider}"`);
  
  let providerFound = false;
  
  if (provider) {
    try {
      await page.getByLabel('Appointment Provider').click();
      await page.waitForTimeout(1000);
      
      try {
        await page.getByRole('menuitem', { name: provider }).click();
        providerFound = true;
        console.log(`✅ Provider selected: ${provider}`);
      } catch (e) {
        console.log('⚠️ Could not select provider, trying alternative strategies...');
        // Try with trailing space
        try {
          await page.getByRole('menuitem', { name: provider + ' ' }).click();
          providerFound = true;
          console.log(`✅ Provider selected (with space): ${provider}`);
        } catch (e2) {
          // Try partial match
          try {
            const allOptions = page.locator('[role="menuitem"]');
            const count = await allOptions.count();
            for (let i = 0; i < count; i++) {
              const option = allOptions.nth(i);
              const text = await option.textContent();
              if (text && text.trim().toLowerCase().includes(provider.toLowerCase())) {
                await option.click();
                providerFound = true;
                console.log(`✅ Provider selected (partial match): "${text.trim()}"`);
                break;
              }
            }
          } catch (e3) {
            console.log('⚠️ Could not select provider with any strategy');
          }
        }
      }
      
      if (!providerFound) {
        console.log(`⚠️ Provider option "${provider}" not found after trying all strategies`);
        // Close dropdown if still open before random click
        try {
          await page.keyboard.press('Escape');
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        try {
          await page.mouse.click(100, 100);
          await page.waitForTimeout(300);
        } catch (e) {
          // Ignore errors
        }
        await page.waitForTimeout(500);
      }
    } catch (e) {
      console.log('⚠️ Error selecting provider:', e.message);
      console.log('⚠️ Could not select provider, skipping');
      // Close dropdown and perform random click
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
  } else {
    console.log('⚠️ No provider data, skipping provider selection');
    providerFound = true; // Consider it as "found" if no data provided
  }
  
  // Select insurance using advanced large dropdown handler
  let insurance = patientData?.['Primary Insurance Name'] || '';
  console.log(`🔍 DEBUG - Insurance data: "${insurance}"`);
  
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
      let optionFound = false;
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
      
      // Strategy 5: Use the advanced scroll and search strategy
      if (!optionFound) {
        console.log('🔍 Trying advanced scroll and search strategy...');
        try {
          const dropdownTrigger = page.getByLabel('Insurance');
          await selectFromLargeDropdown(page, dropdownTrigger, insuranceTrimmed, {
            maxScrollAttempts: 15,
            scrollStep: 200,
            waitTime: 300,
            searchStrategy: 'both'
          });
          optionFound = true;
          console.log(`✅ Insurance selected using advanced dropdown handler: ${insuranceTrimmed}`);
        } catch (e) {
          console.log('🔍 Advanced dropdown handler failed:', e.message);
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
      // Close dropdown and perform random click
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
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed before creating assessment');
    }
    
    await clickAssessmentPlusButton(page);
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed after clicking assessment plus button');
    }
    
    // Verify assessment button clicked successfully
    await verifyAssessmentButtonClicked(page);
    
    await clickCreateAssessment(page);
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed after clicking create assessment');
    }
    
    await selectAssessmentType(page, assessmentType);
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed after selecting assessment type');
    }
    
    // Verify assessment type selected successfully
    await verifyAssessmentTypeSelected(page, assessmentType);
    
    await fillAppointmentForm(page, patientData);
    
    // Check if page is still valid
    if (page.isClosed()) {
      throw new Error('Page was closed after filling appointment form');
    }
    
    // Verify form filled successfully
    await verifyFormFilled(page, patientData);
    
    // Validate form before sending to prevent 422 errors
    console.log('🔍 Validating form before sending...');
    try {
      // Check if all required fields are filled
      const requiredFields = [
        'input[placeholder*="MM/DD/YYYY"]',
        'input[type="text"]',
        'input[type="date"]'
      ];
      
      for (const selector of requiredFields) {
        const fields = page.locator(selector);
        const fieldCount = await fields.count();
        
        for (let i = 0; i < fieldCount; i++) {
          const field = fields.nth(i);
          const isVisible = await field.isVisible();
          if (isVisible) {
            const value = await field.inputValue();
            if (!value || value.trim() === '') {
              console.log(`⚠️ Empty field detected: ${selector} (field ${i})`);
            }
          }
        }
      }
      
      // Check if provider is selected
      const providerField = page.getByLabel('Appointment Provider');
      const providerValue = await providerField.inputValue().catch(() => '');
      if (!providerValue || providerValue.trim() === '') {
        console.log('⚠️ Provider field appears to be empty');
      }
      
      // Check if insurance is selected
      const insuranceField = page.getByLabel('Insurance');
      const insuranceValue = await insuranceField.inputValue().catch(() => '');
      if (!insuranceValue || insuranceValue.trim() === '') {
        console.log('⚠️ Insurance field appears to be empty');
      }
      
      console.log('✅ Form validation completed');
    } catch (validationError) {
      console.log('⚠️ Form validation failed:', validationError.message);
    }
    
    console.log('📤 Sending assessment...');
    
    // Add error handling for the send button click
    try {
      await page.locator('button').filter({ hasText: /^Send$/ }).click();
      console.log('✅ Assessment sent');
      
      // Verify assessment sent successfully
      await verifyAssessmentSent(page);
    } catch (sendError) {
      console.log(`❌ Error clicking send button: ${sendError.message}`);
      throw sendError;
    }
    
    console.log('🔍 DEBUG: About to start popup detection...');
    
    // IMMEDIATELY after sending, detect and handle popup BEFORE any other code
    console.log('🔍 IMMEDIATE: Checking for popup after send button...');
    
    // Wait a moment for popup to appear
    await page.waitForTimeout(2000);
    console.log('🔍 DEBUG: Waited 2 seconds for popup to appear...');
    
    // Take screenshot to see what's on the page
    try {
      console.log('📸 Taking screenshot after send button...');
      await page.screenshot({ path: 'after-send-button.png' });
      console.log('✅ Screenshot saved as after-send-button.png');
    } catch (screenshotError) {
      console.log('⚠️ Screenshot failed:', screenshotError.message);
    }
    
    // Check if page is still valid
    if (page.isClosed()) {
      console.log('❌ Page was closed after sending - cannot detect popup');
      throw new Error('Page was closed after sending assessment');
    }
    
    // Try to detect and handle popup using multiple strategies
    let popupHandled = false;
    
    try {
      // Strategy 1: Look for the specific label using the exact code you provided
      console.log('🔍 Strategy 1: Looking for "This assessment was ordered" label...');
      const popupLabel = page.getByLabel('This assessment was ordered');
      const isPopupVisible = await popupLabel.isVisible({ timeout: 3000 });
      
      if (isPopupVisible) {
        console.log('🔵 CRITICAL: Optional popup detected (Strategy 1), handling it...');
        
        // Check the checkbox using the exact code you provided
        await popupLabel.check();
        console.log('✅ Checked "This assessment was ordered" checkbox');
        
        // Wait a moment for the checkbox to be processed
        await page.waitForTimeout(500);
        
        // Click "Confirm and Continue" button
        const confirmButton = page.locator('button').filter({ hasText: 'Confirm and Continue' });
        await confirmButton.click();
        console.log('✅ Clicked "Confirm and Continue" button');
        
        popupHandled = true;
      } else {
        console.log('🔍 Strategy 1: No popup found with "This assessment was ordered" label');
      }
    } catch (e) {
      console.log('🔍 Strategy 1 failed:', e.message);
    }
    
    // Strategy 2: Look for any checkbox with "assessment was ordered" text
    if (!popupHandled) {
      try {
        console.log('🔍 Strategy 2: Looking for checkbox with "assessment was ordered" text...');
        const checkboxes = page.locator('input[type="checkbox"]');
        const checkboxCount = await checkboxes.count();
        
        console.log(`🔍 Found ${checkboxCount} checkboxes on page`);
        
        for (let i = 0; i < checkboxCount; i++) {
          const checkbox = checkboxes.nth(i);
          const isVisible = await checkbox.isVisible();
          
          if (isVisible) {
            // Check if this checkbox is associated with "assessment was ordered" text
            const parentText = await checkbox.locator('..').textContent();
            const labelText = await checkbox.locator('..').locator('label').textContent().catch(() => '');
            const ariaLabel = await checkbox.getAttribute('aria-label');
            
            console.log(`🔍 Checkbox ${i}: parentText="${parentText}", labelText="${labelText}", ariaLabel="${ariaLabel}"`);
            
            if ((parentText && parentText.toLowerCase().includes('assessment') && parentText.toLowerCase().includes('ordered')) ||
                (labelText && labelText.toLowerCase().includes('assessment') && labelText.toLowerCase().includes('ordered')) ||
                (ariaLabel && ariaLabel.toLowerCase().includes('assessment') && ariaLabel.toLowerCase().includes('ordered'))) {
              
              console.log('🔵 CRITICAL: Optional popup detected (Strategy 2), handling it...');
              
              await checkbox.check();
              console.log('✅ Checked assessment ordered checkbox');
              
              await page.waitForTimeout(500);
              
              const confirmButton = page.locator('button').filter({ hasText: 'Confirm and Continue' });
              await confirmButton.click();
              console.log('✅ Clicked "Confirm and Continue" button');
              
              popupHandled = true;
              break;
            }
          }
        }
      } catch (e) {
        console.log('🔍 Strategy 2 failed:', e.message);
      }
    }
    
    // Strategy 3: Look for any modal/dialog that might contain the popup
    if (!popupHandled) {
      try {
        console.log('🔍 Strategy 3: Looking for modal/dialog...');
        const modal = page.locator('[role="dialog"], .MuiDialog-root, .MuiModal-root, .MuiDialog-container').first();
        const isModalVisible = await modal.isVisible({ timeout: 2000 });
        
        if (isModalVisible) {
          console.log('🔵 CRITICAL: Modal detected (Strategy 3), checking for assessment popup...');
          
          // Look for checkbox inside the modal
          const checkboxes = modal.locator('input[type="checkbox"]');
          const checkboxCount = await checkboxes.count();
          
          console.log(`🔍 Found ${checkboxCount} checkboxes in modal`);
          
          for (let i = 0; i < checkboxCount; i++) {
            const checkbox = checkboxes.nth(i);
            const isVisible = await checkbox.isVisible();
            
            if (isVisible) {
              const parentText = await checkbox.locator('..').textContent();
              const labelText = await checkbox.locator('..').locator('label').textContent().catch(() => '');
              
              console.log(`🔍 Modal checkbox ${i}: parentText="${parentText}", labelText="${labelText}"`);
              
              if ((parentText && parentText.toLowerCase().includes('assessment') && parentText.toLowerCase().includes('ordered')) ||
                  (labelText && labelText.toLowerCase().includes('assessment') && labelText.toLowerCase().includes('ordered'))) {
                
                await checkbox.check();
                console.log('✅ Checked checkbox in modal');
                
                await page.waitForTimeout(500);
                
                const confirmButton = modal.locator('button').filter({ hasText: /confirm.*continue/i });
                await confirmButton.click();
                console.log('✅ Clicked confirm button in modal');
                
                popupHandled = true;
                break;
              }
            }
          }
        } else {
          console.log('🔍 Strategy 3: No modal found');
        }
      } catch (e) {
        console.log('🔍 Strategy 3 failed:', e.message);
      }
    }
    
    if (popupHandled) {
      console.log('✅ Optional popup handled successfully - flow unblocked for next patient');
      
      // Take screenshot after popup handling
      try {
        console.log('📸 Taking screenshot after popup handling...');
        await page.screenshot({ path: 'after-popup-handled.png' });
        console.log('✅ Screenshot saved as after-popup-handled.png');
      } catch (screenshotError) {
        console.log('⚠️ Screenshot failed:', screenshotError.message);
      }
    } else {
      console.log('ℹ️ No optional popup appeared, continuing with normal flow');
      
      // Take screenshot when no popup found
      try {
        console.log('📸 Taking screenshot - no popup found...');
        await page.screenshot({ path: 'no-popup-found.png' });
        console.log('✅ Screenshot saved as no-popup-found.png');
      } catch (screenshotError) {
        console.log('⚠️ Screenshot failed:', screenshotError.message);
      }
    }
    
    // Wait for any immediate response from the server
    await page.waitForTimeout(1500);
    
    // Check for any server errors that might have occurred
    try {
      // Look for error messages on the page
      const errorMessages = page.locator('[role="alert"], .error, .MuiAlert-root, .MuiSnackbar-root');
      const errorCount = await errorMessages.count();
      
      if (errorCount > 0) {
        for (let i = 0; i < errorCount; i++) {
          const errorElement = errorMessages.nth(i);
          const isVisible = await errorElement.isVisible();
          if (isVisible) {
            const errorText = await errorElement.textContent();
            console.log(`⚠️ Server error detected: ${errorText}`);
          }
        }
      }
    } catch (errorCheckError) {
      // Ignore error checking errors
      console.log('🔍 No server errors detected');
    }
    
    
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

async function processPatient(page, dob, apptStr, provider, insurance, assessmentType, isHealthAssessment, dateStr, accountName, customId, firstName, lastName, patientName = '', appointmentDate = null) {
  try {
    console.log(`\n👤 Processing patient DOB: ${dob} (Account: ${accountName})`);
    
    // Check if page is still valid before starting
    if (page.isClosed()) {
      throw new Error('Page was closed before processing patient');
    }
    
    if (!dob || dob === '') {
      console.log(`⚠️ Patient has no DOB - need to add demo`);
      return CONFIG.STATUS.NEED_DEMO;
    }
    
    // Select account first
    await selectAccountOnce(page, accountName, customId);
    
    // Search for DOB immediately after account selection
    console.log(`🔍 Searching for DOB: ${dob}`);
    try {
      // Check if page is still valid
      if (page.isClosed()) {
        throw new Error('Page was closed during DOB search');
      }
      
      // Click on DOB filter field using the specific locator
      await page.getByRole('cell', { name: 'Date of Birth Sort by Date of' }).getByRole('textbox').click();
      
      // Check if page is still valid
      if (page.isClosed()) {
        throw new Error('Page was closed after clicking DOB filter');
      }
      
      // Fill in the DOB
      const dobInput = page.getByRole('textbox', { name: 'MM/DD/YYYY' });
      await dobInput.fill(dob);
      
      console.log(`✅ DOB filter filled with: ${dob}`);
      
      // Click Apply button to trigger search
      await page.getByRole('button', { name: 'Apply' }).click();
      console.log('✅ Clicked Apply button to search');
      
      // Wait for results to load
      await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
      
      // Check if page is still valid
      if (page.isClosed()) {
        throw new Error('Page was closed while waiting for results');
      }
      
      // Check if patient exists by looking for table rows
      const tableRows = page.locator('tbody tr');
      const rowCount = await tableRows.count();
      
      if (rowCount === 0) {
        console.log(`❌ Patient DOB ${dob} not found in system`);
        return CONFIG.STATUS.PATIENT_NOT_FOUND;
      }
      
      console.log(`📊 Found ${rowCount} patient result(s) for DOB ${dob}`);
      
      // If first name or last name is provided, match by name
      // Use separate First Name and Last Name from Excel, fallback to parsing patientName if needed
      let excelFirstName = firstName ? firstName.toLowerCase().trim() : '';
      let excelLastName = lastName ? lastName.toLowerCase().trim() : '';
      
      // Fallback: if separate columns not available, parse from combined patientName
      if (!excelFirstName && !excelLastName && patientName && patientName.trim() !== '') {
        console.log(`🔍 Separate First/Last Name columns not found, parsing from Patient Name: "${patientName}"`);
        const nameParts = patientName.trim().split(/\s+/).filter(part => part.length > 0);
        if (nameParts.length >= 2) {
          excelFirstName = nameParts[0].toLowerCase();
          excelLastName = nameParts[nameParts.length - 1].toLowerCase();
        } else if (nameParts.length === 1) {
          excelFirstName = nameParts[0].toLowerCase();
          excelLastName = nameParts[0].toLowerCase();
        }
      }
      
      if (excelFirstName || excelLastName) {
        console.log(`🔍 Matching patient by name from Excel columns:`);
        console.log(`   First Name: "${excelFirstName || '(not provided)'}"`);
        console.log(`   Last Name: "${excelLastName || '(not provided)'}"`);
        
        let bestMatch = null;
        let bestMatchType = null; // 'exact', 'partial', or null
        
        // Check all rows for name matching
        for (let i = 0; i < rowCount; i++) {
          const row = tableRows.nth(i);
          const rowText = await row.textContent();
          
          if (!rowText) continue;
          
          const lowerRowText = rowText.toLowerCase();
          
          // Check for exact match: both first name and last name found
          let hasFirstName = false;
          let hasLastName = false;
          
          if (excelFirstName) {
            hasFirstName = lowerRowText.includes(excelFirstName);
          }
          if (excelLastName) {
            hasLastName = lowerRowText.includes(excelLastName);
          }
          
          // Exact match: both names found (if both were provided)
          if (excelFirstName && excelLastName && hasFirstName && hasLastName) {
            // Exact match found - use this row
            bestMatch = row;
            bestMatchType = 'exact';
            console.log(`✅ Row ${i + 1}: EXACT MATCH - Found both first name ("${excelFirstName}") and last name ("${excelLastName}"): "${rowText.substring(0, 80)}..."`);
            break; // Stop searching, exact match found
          } else if ((hasFirstName || hasLastName) && !bestMatch) {
            // Partial match - at least one name found, but no exact match yet
            bestMatch = row;
            bestMatchType = 'partial';
            const matchDetails = [];
            if (hasFirstName) matchDetails.push(`first name ("${excelFirstName}")`);
            if (hasLastName) matchDetails.push(`last name ("${excelLastName}")`);
            console.log(`🔍 Row ${i + 1}: PARTIAL MATCH - Found ${matchDetails.join(' and ')}: "${rowText.substring(0, 80)}..."`);
          }
        }
        
        // If no match found (neither exact nor partial), return patient not found
        if (!bestMatch) {
          console.log(`❌ No name match found for patient`);
          console.log(`   Searched ${rowCount} row(s) but could not find matching:`);
          if (excelFirstName) console.log(`     - First Name: "${excelFirstName}"`);
          if (excelLastName) console.log(`     - Last Name: "${excelLastName}"`);
          return CONFIG.STATUS.PATIENT_NOT_FOUND;
        }
        
        // Use the matched row
        const selectedRow = bestMatch;
        console.log(`✅ Selected patient row with ${bestMatchType} name match`);
        
        // Click on ISP cell in the selected row
        const ispCellInRow = selectedRow.locator('td').first();
        
        console.log(`🔍 Clicking ISP cell in matching row...`);
        
        // Take screenshot before clicking
        try {
          await page.screenshot({ path: `before-click-${dob.replace(/\//g, '-')}.png` });
          console.log('📸 Screenshot saved before clicking');
        } catch (e) {}
        
        // Click the ISP cell
        await ispCellInRow.click();
        
        // Wait for navigation to the patient page
        console.log('⏳ Waiting for patient page to load...');
        await page.waitForTimeout(3000);
        
        // Take screenshot after clicking
        try {
          await page.screenshot({ path: `after-click-${dob.replace(/\//g, '-')}.png` });
          console.log('📸 Screenshot saved after clicking');
        } catch (e) {}
        
        console.log(`✅ Clicked ISP cell in matching row: ${dob} (${bestMatchType} match)`);
        
        // Wait for patient dashboard to load
        await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
        
        // Check if page is still valid after patient selection
        if (page.isClosed()) {
          throw new Error('Page was closed after selecting patient');
        }
        
        // Verify patient selection completed successfully
        await verifyPatientSelected(page, dob);
        
      } else {
        // No patient name provided - use first row
        console.log(`⚠️ No patient name provided, selecting first available row`);
        const selectedRow = tableRows.first();
        
        // Click on ISP cell in the selected row
        const ispCellInRow = selectedRow.locator('td').first();
        
        console.log(`🔍 Clicking ISP cell in first row...`);
        
        // Click the ISP cell
        await ispCellInRow.click();
        
        // Wait for navigation to the patient page
        console.log('⏳ Waiting for patient page to load...');
        await page.waitForTimeout(3000);
        
        console.log(`✅ Clicked ISP cell in first row: ${dob}`);
        
        // Wait for patient dashboard to load
        await page.waitForTimeout(CONFIG.TIMEOUTS.MEDIUM);
        
        // Check if page is still valid after patient selection
        if (page.isClosed()) {
          throw new Error('Page was closed after selecting patient');
        }
        
        // Verify patient selection completed successfully
        await verifyPatientSelected(page, dob);
      }
      
    } catch (error) {
      console.log(`❌ Error selecting patient by DOB: ${error.message}`);
      
      // Check if it's a "patient not found" error
      if (error.message.includes('not found') || error.message.includes('DOB')) {
        console.log(`⚠️ Patient DOB ${dob} not found in system - need to add demo`);
        return CONFIG.STATUS.NEED_DEMO;
      }
      
      // Check if it's a page closed error
      if (error.message.includes('Page was closed') || error.message.includes('Target page, context or browser has been closed')) {
        console.log(`❌ Page was closed during patient selection - cannot continue`);
        throw new Error('Page was closed during patient selection');
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
      'Primary Insurance Name': insurance,
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

async function loadExcelData(filePath) {
  const data = await readExcel(filePath);
  return {
    gad16: data['GAD 16'] || [],
    health: data['Health assessment'] || [],
    appointments: data['page'] || []
  };
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

async function processAllPatients(filePath, options = {}) {
  const config = { ...CONFIG, ...options };
  const now = new Date();
  const dateStr = `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
  
  // Launch browser once for all patients
  console.log('🚀 Launching browser for all patients...');
  const browser = await chromium.launch({
    headless: config.HEADLESS,
    slowMo: config.SLOW_MO,
    args: getBrowserArgs()
  });
  
  const context = await browser.newContext({ 
    viewport: { width: 1366, height: 768 },
    ignoreHTTPSErrors: true,
    reducedMotion: 'reduce',
    forcedColors: 'none',
    colorScheme: 'light'
  });
  
  const page = await context.newPage();
  setupPageListeners(page);

  try {
    // Login once for all patients
    await login(page);
    await openSearch(page);
    
    // Load Excel data
    const data = await loadExcelData(filePath);
    
    // Combine all patients into a single array for processing
    const allPatients = [];
    
    // Add GAD16 patients
    data.gad16.forEach(row => {
      allPatients.push({
        ...row,
        type: 'GAD16',
        sheetName: 'GAD 16'
      });
    });
    
    // Add Health Assessment patients
    data.health.forEach(row => {
      allPatients.push({
        ...row,
        type: 'HEALTH',
        sheetName: 'Health assessment'
      });
    });
    
    console.log(`\n📄 Processing ${allPatients.length} total patients (${data.gad16.length} GAD16 + ${data.health.length} Health Assessment)...`);
    
    const results = [];
    
    // Process all patients in a single browser session
    for (let i = 0; i < allPatients.length; i++) {
      const row = allPatients[i];
      const dobRaw = row['MRN'] || ''; // MRN column contains DOB values
      const dob = excelDateToMMDDYYYY(dobRaw); // Convert to MM/DD/YYYY format
      const provider = String(row['Appointment Provider Name'] || row[' Provider'] || row['Scheduler'] || '').trim();
      
      // Find insurance key dynamically (handles special characters)
      const insuranceKey = Object.keys(row).find(key => 
        key.toLowerCase().includes('insurance')
      );
      const insurance = insuranceKey ? String(row[insuranceKey] || '').trim() : '';
      
      // Get account name
      const accountName = String(row['Appointment Facility Name'] || '').trim();
      
      // Get Custom ID - use the correct column name from Excel
      const customId = String(row['Custom ID'] || '').trim();
      
      // Get patient names - use separate First Name and Last Name columns from Excel
      const firstName = String(row['First Name'] || row['FirstName'] || '').trim();
      const lastName = String(row['Last Name'] || row['LastName'] || row['Last Name'] || '').trim();
      
      // Fallback: try combined Patient Name column if separate columns not found
      const patientName = String(row['Patient Name'] || row['Name'] || row['Patient Name (Last, First)'] || row['Full Name'] || '').trim();
      
      // Get appointment date from Excel - try different possible column names
      const appointmentDateRaw = row['Appointment Date'] || row['AppointmentDate'] || row['Appt Date'] || '';
      const appointmentDate = appointmentDateRaw ? excelDateToMMDDYYYY(appointmentDateRaw) : null;
      
      console.log(`\n🔍 Processing patient ${i + 1}/${allPatients.length} - DOB: ${dob} (Type: ${row.type})`);
      console.log(`  Provider: "${provider}"`);
      console.log(`  Insurance: "${insurance}"`);
      console.log(`  Account: "${accountName}"`);
      console.log(`  Custom ID: "${customId}"`);
      console.log(`  First Name: "${firstName}"`);
      console.log(`  Last Name: "${lastName}"`);
      console.log(`  Patient Name (fallback): "${patientName}"`);
      console.log(`  Appointment Date: "${appointmentDate || 'Not specified (will use default)'}"`);
      
      try {
        // Determine assessment type and parameters
        const assessmentType = row.type === 'GAD16' ? CONFIG.ASSESSMENT_TYPES.PHQ_GAD16 : CONFIG.ASSESSMENT_TYPES.HEALTH;
        const isHealthAssessment = row.type === 'HEALTH';
        
        const result = await processPatient(page, dob, null, provider, insurance, assessmentType, isHealthAssessment, dateStr, accountName, customId, firstName, lastName, patientName, appointmentDate);
        const status = result.result || result;
        row.Status = status;

        // Decide best identifier to update the correct row
        // Priority: MRN (DOB) -> First Name + Last Name -> Patient Name -> Custom ID
        // MRN is most reliable since each patient has unique DOB
        let searchColumn = 'MRN';
        let searchValue = excelDateToMMDDYYYY(row['MRN']);
        
        // If MRN is not available, try First Name + Last Name combination
        if (!searchValue || !row['MRN']) {
          if (firstName && lastName) {
            // Try to match using First Name and Last Name combination
            searchColumn = 'First Name';
            searchValue = firstName;
            // Note: We'll need to match both First Name and Last Name when updating
          } else if (patientName) {
            searchColumn = 'Patient Name';
            searchValue = patientName;
          } else if (customId) {
            // Only use Custom ID as last resort
            searchColumn = 'Custom ID';
            searchValue = customId;
          }
        }

        console.log(`\n💾 Saving status to Excel: ${row.sheetName} - ${searchColumn}: ${searchValue} -> Status: ${status}`);
        
        // Update status with retry logic
        let excelUpdated = false;
        
        // If using First Name, try to match with both First Name and Last Name
        if (searchColumn === 'First Name' && firstName && lastName) {
          // Try to find row matching both First Name and Last Name
          excelUpdated = await markExcel(filePath, row.sheetName, 'First Name', firstName, status, 'Last Name', lastName);
          if (!excelUpdated) {
            // Fallback: try with just First Name
            excelUpdated = await markExcel(filePath, row.sheetName, 'First Name', firstName, status);
          }
        } else {
          excelUpdated = await markExcel(filePath, row.sheetName, searchColumn, searchValue, status);
        }
        
        // If primary search fails, try alternative columns
        if (!excelUpdated) {
          if (searchColumn !== 'MRN' && row['MRN']) {
            // Retry with MRN if other methods failed
            console.log(`🔄 Retrying status update with MRN...`);
            const dobValue = excelDateToMMDDYYYY(row['MRN']);
            excelUpdated = await markExcel(filePath, row.sheetName, 'MRN', dobValue, status);
            if (excelUpdated) {
              searchColumn = 'MRN';
              searchValue = dobValue;
              console.log(`✅ Status updated using MRN: ${dobValue}`);
            }
          } else if (searchColumn !== 'First Name' && firstName && lastName) {
            // Retry with First Name + Last Name
            console.log(`🔄 Retrying status update with First Name + Last Name...`);
            excelUpdated = await markExcel(filePath, row.sheetName, 'First Name', firstName, status, 'Last Name', lastName);
            if (excelUpdated) {
              searchColumn = 'First Name';
              searchValue = firstName;
              console.log(`✅ Status updated using First Name + Last Name`);
            }
          } else if (searchColumn !== 'Patient Name' && patientName) {
            // Retry with Patient Name if MRN failed
            console.log(`🔄 Retrying status update with Patient Name...`);
            excelUpdated = await markExcel(filePath, row.sheetName, 'Patient Name', patientName, status);
            if (excelUpdated) {
              searchColumn = 'Patient Name';
              searchValue = patientName;
              console.log(`✅ Status updated using Patient Name: ${patientName}`);
            }
          }
        }
        
        if (!excelUpdated) {
          console.error(`❌ Failed to update Excel status for ${searchColumn}: ${searchValue}`);
          const patientInfo = [];
          if (firstName) patientInfo.push(`First: ${firstName}`);
          if (lastName) patientInfo.push(`Last: ${lastName}`);
          if (patientName) patientInfo.push(`Full: ${patientName}`);
          console.error(`   Patient: ${patientInfo.join(', ') || 'N/A'}, DOB: ${dob}, Custom ID: ${customId || 'N/A'}`);
        } else {
          console.log(`✅ Successfully updated status for patient ${i + 1}/${allPatients.length}`);
          // Small delay to ensure file is fully written (especially important for OneDrive sync)
          await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        // Update Last Order Date if status is SENT
        if (status === CONFIG.STATUS.SENT) {
          let dateUpdated = false;
          
          // If using First Name, try to match with both First Name and Last Name
          if (searchColumn === 'First Name' && firstName && lastName) {
            dateUpdated = await markExcel(filePath, row.sheetName, 'First Name', firstName, dateStr, 'Last Order Date', 'Last Name', lastName);
            if (!dateUpdated) {
              dateUpdated = await markExcel(filePath, row.sheetName, 'First Name', firstName, dateStr, 'Last Order Date');
            }
          } else {
            dateUpdated = await markExcel(filePath, row.sheetName, searchColumn, searchValue, dateStr, 'Last Order Date');
          }
          
          if (!dateUpdated) {
            console.error(`❌ Failed to update Last Order Date for ${searchColumn}: ${searchValue}`);
            // Retry with alternative columns if available
            if (searchColumn !== 'MRN' && row['MRN']) {
              const dobValue = excelDateToMMDDYYYY(row['MRN']);
              console.log(`🔄 Retrying Last Order Date update with MRN...`);
              await markExcel(filePath, row.sheetName, 'MRN', dobValue, dateStr, 'Last Order Date');
            } else if (searchColumn !== 'First Name' && firstName && lastName) {
              console.log(`🔄 Retrying Last Order Date update with First Name + Last Name...`);
              await markExcel(filePath, row.sheetName, 'First Name', firstName, dateStr, 'Last Order Date', 'Last Name', lastName);
            } else if (searchColumn !== 'Patient Name' && patientName) {
              console.log(`🔄 Retrying Last Order Date update with Patient Name...`);
              await markExcel(filePath, row.sheetName, 'Patient Name', patientName, dateStr, 'Last Order Date');
            }
          } else {
            // Small delay to ensure file is fully written
            await new Promise(resolve => setTimeout(resolve, 500));
          }
        }
        
        results.push({
          dob,
          type: row.type,
          status,
          accountName,
          customId
        });
        
        // Navigate back to accounts page after each assessment (except for the last patient)
        if (i < allPatients.length - 1) {
          await navigateBackToAccounts(page);
        }
        
      } catch (error) {
        console.error(`❌ Error processing patient ${dob}:`, error.message);
        
        // Check if this is a verification failure - if so, stop the process
        if (error.message.includes('verification failed') || error.message.includes('Verification failed')) {
          console.error(`\n🛑 CRITICAL: Step verification failed - stopping process to prevent incomplete operations`);
          console.error(`   Error: ${error.message}`);
          console.error(`   Patient: ${patientName || 'N/A'}, DOB: ${dob}, Account: ${accountName}`);
          console.error(`   Processed ${i}/${allPatients.length} patients before stopping`);
          
          // Mark current patient as error
          row.Status = CONFIG.STATUS.ERROR;
          let searchColumn = 'MRN';
          let searchValue = excelDateToMMDDYYYY(row['MRN']);
          
          if (!searchValue || !row['MRN']) {
            if (patientName) {
              searchColumn = 'Patient Name';
              searchValue = patientName;
            } else if (customId) {
              searchColumn = 'Custom ID';
              searchValue = customId;
            }
          }
          
          if (searchValue) {
            await markExcel(filePath, row.sheetName, searchColumn, searchValue, CONFIG.STATUS.ERROR).catch(() => {});
          }
          
          // Stop processing - throw error to exit loop
          throw new Error(`Process stopped due to verification failure: ${error.message}`);
        }
        
        row.Status = CONFIG.STATUS.ERROR;
        
        // Use the same search logic as success case to find the correct row
        // Priority: MRN (DOB) -> Patient Name -> Custom ID
        let searchColumn = 'MRN';
        let searchValue = excelDateToMMDDYYYY(row['MRN']);
        
        // If MRN is not available, try Patient Name
        if (!searchValue || !row['MRN']) {
          if (patientName) {
            searchColumn = 'Patient Name';
            searchValue = patientName;
          } else if (customId) {
            // Only use Custom ID as last resort
            searchColumn = 'Custom ID';
            searchValue = customId;
          }
        }
        
        // Only try to update if we have a valid search value
        if (searchValue) {
          let errorUpdated = await markExcel(filePath, row.sheetName, searchColumn, searchValue, CONFIG.STATUS.ERROR);
          if (!errorUpdated) {
            console.error(`❌ Failed to update Excel with ERROR status for ${searchColumn}: ${searchValue}`);
            // Try alternative search methods if first attempt failed
            if (searchColumn !== 'MRN' && row['MRN']) {
              const dobValue = excelDateToMMDDYYYY(row['MRN']);
              console.log(`🔄 Retrying with MRN: ${dobValue}`);
              errorUpdated = await markExcel(filePath, row.sheetName, 'MRN', dobValue, CONFIG.STATUS.ERROR);
            } else if (searchColumn !== 'Patient Name' && patientName) {
              console.log(`🔄 Retrying with Patient Name: ${patientName}`);
              errorUpdated = await markExcel(filePath, row.sheetName, 'Patient Name', patientName, CONFIG.STATUS.ERROR);
            }
          }
        } else {
          console.error(`❌ Cannot update Excel - no valid identifier found (Custom ID, MRN, or Patient Name is empty)`);
        }
        
        results.push({
          dob,
          type: row.type,
          status: CONFIG.STATUS.ERROR,
          accountName,
          customId,
          error: error.message
        });
        
        // Try to navigate back to accounts page even after error
        if (i < allPatients.length - 1) {
          try {
            await navigateBackToAccounts(page);
          } catch (navError) {
            console.error('❌ Error navigating back after error:', navError.message);
          }
        }
      }
      
      // Add delay between patients (except for the last one)
      if (i < allPatients.length - 1) {
        console.log('⏳ Waiting before next patient...');
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
    }
    
    console.log('\n🎉 All patients processed successfully!');
    return {
      success: true,
      totalProcessed: allPatients.length,
      results: results
    };
    
  } catch (error) {
    console.error('\n❌ Fatal error during patient processing:', error.message);
    
    // If it's a verification failure, make it clear
    if (error.message.includes('verification failed') || error.message.includes('Verification failed') || error.message.includes('Process stopped')) {
      console.error('\n🛑 PROCESS STOPPED: Step verification failed - incomplete operation prevented');
      console.error('   Please check the error above and fix the issue before continuing');
    }
    
    return {
      success: false,
      error: error.message,
      stopped: error.message.includes('verification failed') || error.message.includes('Verification failed') || error.message.includes('Process stopped')
    };
  } finally {
    // Close browser only after all patients are processed
    console.log('⏳ Closing browser...');
    await new Promise(resolve => setTimeout(resolve, 5000));
    await browser.close();
  }
}

module.exports = {
  processAllPatients,
  loadExcelData
};



