const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    
    if (action === 'login') result = loginUser(data.payload);
    else if (action === 'vendorSignUp') result = vendorSignUp(data.payload);
    else if (action === 'saveInquiry') result = saveInquiry(data.payload);
    else if (action === 'getAdminStats') result = getAdminStats();
    else if (action === 'getAdminVendorsList') result = getAdminVendorsList();
    else if (action === 'approveVendor') result = approveVendor(data.payload);
    else if (action === 'deleteVendor') result = deleteVendor(data.payload);
    else if (action === 'updateVendorLogo') result = updateVendorLogo(data.payload);
    else if (action === 'crudPlan') result = crudPlan(data.payload);
    else if (action === 'getAllLeadsAdmin') result = getAllLeadsAdmin();
    else if (action === 'adminDeleteLead') result = adminDeleteLead(data.payload);
    else if (action === 'getAllInquiriesAdmin') result = getAllInquiriesAdmin();
    else if (action === 'adminDeleteInquiry') result = adminDeleteInquiry(data.payload);
    else if (action === 'getVendorLeads') result = getVendorLeads(data.payload.vendorId);
    else if (action === 'calculateSystem') result = calculateSystem(data.payload);
    else if (action === 'generateQuotation') result = generateQuotation(data.payload);
    else if (action === 'acceptQuotationAndGenerateDocs') result = acceptQuotationAndGenerateDocs(data.payload);
    else if (action === 'updateLeadStage') result = updateLeadStage(data.payload);
    else if (action === 'deleteLead') result = deleteLead(data.payload);
    else if (action === 'crudTask') result = handleTasks(data.payload);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function loginUser(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][2] === payload.email && data[i][4] === payload.password) {
      if(data[i][6] !== 'Approved' && data[i][0] !== 'ADMIN') {
        return { success: false, message: 'Account is pending Master Admin approval. You cannot log in yet.' };
      }
      return { 
        success: true, 
        role: data[i][0] === 'ADMIN' ? 'admin' : 'vendor', 
        vendorId: data[i][0], 
        name: data[i][1],
        logoUrl: data[i][11] || ''
      };
    }
  }
  return { success: false, message: 'Invalid credentials or unregistered email.' };
}

function vendorSignUp(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][2] === payload.email) {
      return { success: false, message: 'Email already registered. Please login.' };
    }
  }
  const vendorId = 'VND' + new Date().getTime();
  sheet.appendRow([vendorId, payload.name, payload.email, payload.phone, payload.password, 'Active', 'Pending', 'Free Tier', '', '', payload.address, '', '', '']);
  return { success: true, message: 'Sign-up successful! Your vendor account is now pending Master Admin approval.' };
}

function saveInquiry(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Inquiries') || SpreadsheetApp.openById(SHEET_ID).insertSheet('Inquiries');
  if(sheet.getLastRow() === 0) sheet.appendRow(['Timestamp', 'Org_Name', 'Name', 'Phone', 'Address', 'Specific_Service', 'Requirements', 'Source']);
  sheet.appendRow([new Date(), payload.orgName || 'N/A', payload.name, payload.phone, payload.address || 'N/A', payload.service, payload.requirements, payload.source || 'AI Chat']);
  return { success: true, message: 'Inquiry saved successfully.' };
}

function getAllInquiriesAdmin() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Inquiries');
  if(!sheet) return { success: true, inquiries: [] };
  const data = sheet.getDataRange().getValues();
  let inquiries = [];
  for(let i=1; i<data.length; i++) {
    inquiries.push({ id: i, time: data[i][0], org: data[i][1], name: data[i][2], phone: data[i][3], address: data[i][4], service: data[i][5], req: data[i][6], source: data[i][7] });
  }
  return { success: true, inquiries: inquiries };
}

function adminDeleteInquiry(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Inquiries');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(i === payload.index) { sheet.deleteRow(i+1); return { success: true, message: 'Inquiry deleted.' }; }
  }
  return { success: false, message: 'Inquiry not found.' };
}

function getAdminStats() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  let pending = 0, active = 0, total = 0;
  for(let i=1; i<data.length; i++) {
    if(data[i][0] !== 'ADMIN') {
      total++;
      if(data[i][6] === 'Pending') pending++;
      if(data[i][6] === 'Approved') active++;
    }
  }
  return { total, active, pending };
}

function getAdminVendorsList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  let vendors = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][0] !== 'ADMIN') {
      vendors.push({ vendorId: data[i][0], name: data[i][1], email: data[i][2], phone: data[i][3], status: data[i][5], approval: data[i][6], plan: data[i][7], address: data[i][10], logoUrl: data[i][11] || '' });
    }
  }
  return vendors;
}

function approveVendor(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.vendorId) { sheet.getRange(i+1, 7).setValue('Approved'); return { success: true, message: 'Vendor approved successfully. They can now log in.' }; }
  }
  return { success: false, message: 'Not found.' };
}

function deleteVendor(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.vendorId) { sheet.deleteRow(i + 1); return { success: true, message: 'Deleted.' }; }
  }
}

function updateVendorLogo(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Vendors');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.vendorId) { 
      sheet.getRange(i+1, 12).setValue(payload.logoUrl);
      return { success: true, message: 'Vendor logo updated successfully.' }; 
    }
  }
  return { success: false, message: 'Vendor not found.' };
}

function crudPlan(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Plans');
  if(payload.method === 'READ') {
    const data = sheet.getDataRange().getValues();
    let plans = [];
    for(let i=1; i<data.length; i++) { plans.push({ id: data[i][0], name: data[i][1], price: data[i][2], features: data[i][3] }); }
    return { success: true, plans: plans };
  } else if(payload.method === 'CREATE') {
    sheet.appendRow(['PLN'+new Date().getTime(), payload.name, payload.price, payload.features]);
    return { success: true, message: 'Plan created.' };
  } else if(payload.method === 'DELETE') {
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === payload.planId) { sheet.deleteRow(i+1); return { success: true, message: 'Plan deleted.' }; }
    }
  }
}

function getAllLeadsAdmin() {
  const leadSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = leadSheet.getDataRange().getValues();
  let leads = [];
  for(let i=1; i<data.length; i++) {
    leads.push({ leadId: data[i][0], vendorId: data[i][1], customer: data[i][2], status: data[i][4], kw: data[i][7] });
  }
  return { success: true, leads: leads };
}

function adminDeleteLead(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.leadId) { sheet.deleteRow(i+1); return { success: true, message: 'Lead deleted by admin.' }; }
  }
}

function getVendorLeads(vendorId) {
  const leadSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = leadSheet.getDataRange().getValues();
  let leads = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][1] === vendorId) {
      leads.push({ 
        leadId: data[i][0], 
        date: data[i][12], 
        customer: data[i][2], 
        phone: data[i][3], 
        status: data[i][4], 
        kw: data[i][7], 
        rate: data[i][8],
        moduleBrand: data[i][9], 
        moduleWp: data[i][10] || 550,
        moduleQty: data[i][11] || 0,
        inverterBrand: data[i][12], 
        totalAmount: data[i][13],
        address: data[i][14], 
        geoLocation: data[i][15] 
      });
    }
  }
  return leads;
}

function calculateSystem(payload) {
  const kwRequired = Math.ceil(((payload.billAmount / 8) / 120) * 10) / 10;
  return { kw: kwRequired, units: Math.round(payload.billAmount / 8) };
}

function generateQuotation(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const leadId = 'LD-' + new Date().getTime();
  const totalAmt = parseFloat(payload.kw) * parseFloat(payload.ratePerKw);
  sheet.appendRow([
    leadId, payload.vendorId, payload.customer, payload.phone, 'Quotation Sent', 'Awaiting Consumer Acceptance', '', 
    payload.kw, payload.ratePerKw, payload.moduleBrand, payload.moduleWp, payload.moduleQty, payload.inverterBrand, 
    totalAmt, new Date().toISOString().split('T')[0], payload.address, payload.geoLocation
  ]);
  return { success: true, leadId: leadId, message: 'Quotation generated.' };
}

function acceptQuotationAndGenerateDocs(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.leadId) { sheet.getRange(i+1, 5).setValue('Quotation Accepted'); return { success: true, message: 'Accepted.' }; }
  }
}

function updateLeadStage(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.leadId) { sheet.getRange(i+1, 5).setValue(payload.newStatus); return { success: true, message: 'Updated.' }; }
  }
}

function deleteLead(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leads');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === payload.leadId) { sheet.deleteRow(i+1); return { success: true, message: 'Deleted.' }; }
  }
}

function handleTasks(payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Tasks');
  if(payload.method === 'CREATE') {
    sheet.appendRow(['TSK'+new Date().getTime(), payload.vendorId, payload.leadId, payload.taskName, 'Pending', payload.dueDate]);
    return { success: true, message: 'Task added.' };
  } else if (payload.method === 'READ') {
    const data = sheet.getDataRange().getValues();
    let tasks = [];
    for(let i=1; i<data.length; i++) {
      if(data[i][1] === payload.vendorId) { tasks.push({ id: data[i][0], leadId: data[i][2], name: data[i][3], status: data[i][4], due: data[i][5] }); }
    }
    return { success: true, tasks: tasks };
  } else if (payload.method === 'DELETE') {
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === payload.taskId) { sheet.deleteRow(i+1); return { success: true, message: 'Task deleted.' }; }
    }
  }
}
