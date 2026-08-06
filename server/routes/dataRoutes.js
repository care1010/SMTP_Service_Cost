const express = require('express');
const router = express.Router();
const dataController = require('../controllers/dataController');
const ptdController = require('../controllers/ptdController');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' });
const projectController = require('../controllers/projectController');
const path = require('path');
// ASBL Update route
const asblController = require('../controllers/asblController');
const authController = require('../controllers/authController');
const { getDropdownData } = require('../controllers/accessRequestController'); // Import the new controller function
const rawDrillController = require('../controllers/rawDrillController');


// In teeno routes ke naam controller ke function se match hone chahiye
router.get('/wbs-summary', dataController.getWbsSummary);
router.get(
'/wbs-summary-collapse', dataController.getWbsSummaryCollapse);

router.get('/filter-options', dataController.getFilterOptions); // Is line ki wajah se crash ho raha tha
router.get('/export-excel', dataController.exportToExcel);
router.get('/categories', dataController.getCategories);
router.post('/save-project', dataController.saveProjectData);
router.post('/update-non-committed', dataController.updateNonCommitted);
router.post('/ptd-automation', upload.single('file'), ptdController.uploadPtdData);
router.get('/review-changes', dataController.getReviewChanges);
router.post('/finalize-changes', dataController.finalizeChanges);
router.post('/process-project-paste', projectController.processProjectPaste);
router.get('/download-template', (req, res) => {
    const filePath = path.join(__dirname, '../../ASBL_Data_Template.xlsx');
    res.download(filePath);
});

router.post('/login', authController.login);

router.get('/dashboard-filters', dataController.getDashboardFilters);

router.get('/analytics-bu', dataController.getBuAnalytics);

router.get('/analytics-loa', dataController.getLoaAnalytics);

router.post('/full-refresh', dataController.fullRefresh);

// Project Template download karne ka route
router.get('/download-project-template', (req, res) => {
    // Aapke bataye huye path ke hisaab se
    const filePath = path.join(__dirname, '../../Add_New_Project.xlsx');
    res.download(filePath, 'Add_New_Project.xlsx', (err) => {
        if (err) {
            res.status(500).send("Template file not found on server.");
        }
    });
});

router.get('/project-details', asblController.getProjectDetails);
router.post('/update-manual-asbl', asblController.updateManualAsbl);

// ASBL Update route
router.post('/process-asbl-update', asblController.processAsblUpdate);
// ASBL Template download karne ka route
router.get('/download-asbl-template', (req, res) => {
    const filePath = path.join(__dirname, '../../ASBL_Data_Template.xlsx');
    res.download(filePath, 'ASBL_Data_Template.xlsx', (err) => {
        if (err) res.status(500).send("ASBL Template not found.");
    });
});

router.get('/download-ptd-template', (req, res) => {
    const filePath = path.join(__dirname, '../../PTD_template.xlsx');
    res.download(filePath, 'PTD_template.xlsx', (err) => {
        if (err) res.status(500).send("PTD Template not found.");
    });
});


const adminController = require('../controllers/adminController');
router.get('/admin/users', adminController.getAllUsers);
router.post('/admin/create-user', adminController.createUser);
router.post('/admin/update-user', adminController.updateUser);
router.delete('/admin/delete-user', adminController.deleteUser);

router.post('/clear-draft', dataController.clearDraftChanges);
router.get('/export-review', dataController.exportReviewExcel);
router.post('/finalize-changes', dataController.finalizeChanges);

router.get('/check-pending-changes', dataController.checkPendingChanges);

router.get('/final-dashboard-table', dataController.getFinalDashboardTable);

router.get('/cost-view-table', dataController.getCostViewTable);

router.get('/customer-view-table', dataController.getCustomerViewTable);

router.get('/bu-customer-view-table', dataController.getBuCustomerViewTable);

router.get('/customer-bu-view-table', dataController.getCustomerBuViewTable);

router.get('/customer-bu-loa-view-table', dataController.getCustomerBuLoaViewTable);

router.post('/drilldown', dataController.getDrillDownData);
router.get('/export-drilldown', dataController.exportDrillDown);
router.get('/negative-loa-table', dataController.getNegativeLOATable);

router.get("/user-activity-logs", dataController.getUserActivityLogs);

router.get("/pending-users", dataController.getPendingUsers);

router.get('/non-committed-trend', dataController.getNonCommittedTrend);

router.get('/trend-loas', dataController.getTrendLoas);

router.get('/erp-resource', dataController.getERPResource);

router.get('/erp-resource-export', dataController.exportERPResource);

router.post('/erp-resource/upload',  upload.single('file'), dataController.uploadERPResource);

// Example routes mapping
router.post('/upload-project-file', upload.single('file'), projectController.uploadProjectFile);

router.get('/project-wbs-options', asblController.getProjectWbsOptions);

// router.get('/asbl-wbs-types', asblController.getAsblWbsTypes); 
router.get('/filtered-projects', asblController.getFilteredProjects);

// Ye naya route add karein:
router.get('/run-fix-db', projectController.fixMissingSummaryRows);

router.get('/admin/master-customers', adminController.getAllMasterCustomers);

// router.get('/test-ptd-seeding', projectController.testPtdSeeding);

router.get('/add-project-options', projectController.getAddProjectOptions);

// Route to get all dropdown lists
router.get('/dropdowns', getDropdownData);

const accessRequestController = require('../controllers/accessRequestController');
router.get('/access/dropdowns', accessRequestController.getDropdownData);
router.post('/access/request', accessRequestController.submitRequest);
router.get('/access/pending', accessRequestController.getPendingRequests);
router.post('/access/approve', accessRequestController.approveRequest);
router.post('/access/decline', accessRequestController.declineRequest);

// Raw Data Explorer (CJ74 / CJI5) Routes
router.get('/raw-get-data', rawDrillController.getRawData);
router.get('/raw-export', rawDrillController.exportRawData);

module.exports = router;