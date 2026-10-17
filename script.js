/**
 * ==========================================
 * 1. GLOBAL VARIABLES & MOCK DATA
 * ==========================================
 * Note: This section contains placeholder data. 
 * In production, this data should be fetched from SharePoint / MS Graph API.
 */

/**
 * ==========================================
 * 1. AZURE AUTHENTICATION & MS GRAPH INIT
 * ==========================================
 */

// 1. إعدادات الاتصال بـ Azure
const msalConfig = {
    auth: {
        clientId: "05bcce89-b01a-4732-857f-8402f7a649f5", 
        authority: "https://login.microsoftonline.com/58d349d2-2319-41f9-9aff-f3055d0a1880", 
        redirectUri: "https://haider414.github.io/e-signature-portal/" 
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let currentUserEmail = ""; // سنملأ هذا المتغير برمجياً بعد تسجيل الدخول

// 2. دالة تسجيل الدخول التلقائي
async function authenticateUser() {
    try {
        // طلب تسجيل الدخول وقراءة البيانات
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read", "Sites.Read.All"]
        });
        
        // استخراج بيانات الموظف
        const user = loginResponse.account;
        currentUserEmail = user.username; // الإيميل الحقيقي
        const fullName = user.name;       // الاسم الحقيقي
        
        // 3. تحديث الواجهة (HTML) ببيانات الموظف الحقيقية
        document.querySelector('.user-name').textContent = fullName;
        document.getElementById('senderName').value = fullName;
        document.getElementById('senderEmail').value = currentUserEmail;
        
        console.log("Authentication Successful! User:", currentUserEmail);
        
        // تحميل بيانات الداشبورد بعد نجاح تسجيل الدخول
        loadDashboardData();
        
    } catch (error) {
        console.error("Authentication Error:", error);
        alert("فشل تسجيل الدخول. يرجى التأكد من صلاحيات Azure.");
    }
}

// تشغيل دالة تسجيل الدخول فور فتح الصفحة
window.addEventListener('load', () => {
    authenticateUser();
});

// Mock database for testing the UI
const mockRequests = [
    // --- Sent Requests (Created by the current user) ---
    {
        id: "REQ-2026-001",
        senderName: "IT Administrator",
        senderEmail: "it_dept@dammamvalley.sa",
        recipientName: "Ahmed Almutairi",
        recipientEmail: "ahmed@dammamvalley.sa",
        fileName: "NDA_Contract.pdf",
        date: "Oct 06, 2026",
        status: "Pending"
    },
    {
        id: "REQ-2026-002",
        senderName: "IT Administrator",
        senderEmail: "it_dept@dammamvalley.sa",
        recipientName: "External Partner",
        recipientEmail: "partner@external.com",
        fileName: "Vendor_Agreement.pdf",
        date: "Oct 08, 2026",
        status: "Signed"
    },
    {
        id: "REQ-2026-003",
        senderName: "IT Administrator",
        senderEmail: "it_dept@dammamvalley.sa",
        recipientName: "Khalid Alshahrani",
        recipientEmail: "khalid@dammamvalley.sa",
        fileName: "IT_Policy_Acknowledgment.pdf",
        date: "Oct 10, 2026",
        status: "Declined"
    },
    
    // --- Received Requests (Sent to the current user) ---
    {
        id: "REQ-2026-004",
        senderName: "HR Department",
        senderEmail: "hr@dammamvalley.sa",
        recipientName: "IT Administrator",
        recipientEmail: "it_dept@dammamvalley.sa",
        fileName: "Annual_Evaluation.docx",
        date: "Oct 09, 2026",
        status: "Pending" 
    },
    {
        id: "REQ-2026-005",
        senderName: "Executive Office",
        senderEmail: "ceo@dammamvalley.sa",
        recipientName: "IT Administrator",
        recipientEmail: "it_dept@dammamvalley.sa",
        fileName: "Budget_Approval_Q4.pdf",
        date: "Oct 01, 2026",
        status: "Signed" 
    }
];


/**
 * ==========================================
 * 2. UI NAVIGATION & STATE MANAGEMENT
 * ==========================================
 * Handles switching between the "New Request" form and the "Dashboard" views.
 */

function showSection(sectionId) {
    // Hide all sections initially
    document.getElementById('newRequestSection').style.display = 'none';
    document.getElementById('dashboardSection').style.display = 'none';
    
    // Remove 'active' class from all sidebar links
    document.querySelectorAll('.nav-links li').forEach(li => li.classList.remove('active'));
    
    // Show the targeted section and highlight the corresponding sidebar link
    if (sectionId === 'newRequest') {
        document.getElementById('newRequestSection').style.display = 'block';
        document.querySelectorAll('.nav-links li')[0].classList.add('active');
    } else if (sectionId === 'dashboard') {
        document.getElementById('dashboardSection').style.display = 'block';
        document.querySelectorAll('.nav-links li')[1].classList.add('active');
        
        // Refresh dashboard data every time the view is opened
        loadDashboardData(); 
    }
}


/**
 * ==========================================
 * 3. DASHBOARD RENDERING (Sent/Received)
 * ==========================================
 * Filters the data and populates the HTML tables.
 */

function loadDashboardData() {
    const sentBody = document.getElementById('sentRequestsBody');
    const receivedBody = document.getElementById('receivedRequestsBody');
    
    // Clear existing table rows to prevent duplication
    sentBody.innerHTML = ''; 
    receivedBody.innerHTML = ''; 

    // Loop through the data and append rows based on user role (Sender vs. Recipient)
    mockRequests.forEach((request, index) => {
        // Determine the CSS class for the status badge
        let statusClass = 'pending';
        if (request.status === 'Signed') statusClass = 'signed';
        if (request.status === 'Declined') statusClass = 'declined';
        
        let tr = document.createElement('tr');
        
        // Condition 1: Current user is the Sender
        if (request.senderEmail === currentUserEmail) {
            tr.innerHTML = `
                <td>${request.recipientName}</td>
                <td>${request.fileName}</td>
                <td>${request.date}</td>
                <td><span class="status ${statusClass}">${request.status}</span></td>
                <td><button class="btn-view" onclick="openDynamicModal(${index})">View Details</button></td>
            `;
            sentBody.appendChild(tr);
        } 
        // Condition 2: Current user is the Recipient
        else if (request.recipientEmail === currentUserEmail) {
            tr.innerHTML = `
                <td>${request.senderName}</td>
                <td>${request.fileName}</td>
                <td>${request.date}</td>
                <td><span class="status ${statusClass}">${request.status}</span></td>
                <td><button class="btn-view" onclick="openDynamicModal(${index})">View Details</button></td>
            `;
            receivedBody.appendChild(tr);
        }
    });
}

// Initialize the dashboard data when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', loadDashboardData);


/**
 * ==========================================
 * 4. MODAL MANAGEMENT (Details Popup)
 * ==========================================
 * Handles opening/closing the details modal and injecting specific request data.
 */

function openDynamicModal(index) {
    const request = mockRequests[index];
    
    // Inject specific request data into the modal's HTML elements
    document.getElementById('modalSenderName').textContent = request.senderName;
    document.getElementById('modalSenderEmail').textContent = request.senderEmail;
    
    document.getElementById('modalRecipientName').textContent = request.recipientName;
    document.getElementById('modalRecipientEmail').textContent = request.recipientEmail;
    
    document.getElementById('modalFileName').textContent = request.fileName;
    document.getElementById('modalOrderId').textContent = request.id;
    document.getElementById('modalDate').textContent = request.date;
    
    // Update the visual status badge in the modal header
    const statusBadge = document.getElementById('modalStatusBadge');
    statusBadge.textContent = request.status;
    
    if (request.status === 'Signed') {
        statusBadge.className = 'status signed';
    } else if (request.status === 'Declined') {
        statusBadge.className = 'status declined'; 
    } else {
        statusBadge.className = 'status pending';
    }
    
    // Reminder Button Logic: 
    // Only show if the status is "Pending" AND the current user is the original sender
    const btnReminder = document.getElementById('btnSendReminder');
    if (request.status === 'Pending' && request.senderEmail === currentUserEmail) {
        btnReminder.style.display = 'block'; 
    } else {
        btnReminder.style.display = 'none';  
    }
    // (View Document)
    const btnView = document.getElementById('btnViewDocument');
    btnView.onclick = function() {
        alert("This feature will be connected to SharePoint soon to view: " + request.fileName);
    };
    // Display the modal overlay
    document.getElementById('detailsModal').style.display = 'block';
}

function closeModal() {
    document.getElementById('detailsModal').style.display = 'none';
}

// Close the modal when the user clicks anywhere outside of the modal content
window.onclick = function(event) {
    const modal = document.getElementById('detailsModal');
    if (event.target == modal) {
        modal.style.display = "none";
    }
}


/**
 * ==========================================
 * 5. FORM SUBMISSION & API INTEGRATION
 * ==========================================
 * Handles the file upload UI, Base64 conversion, and POST request to Power Automate.
 */

// Handle File Input UI (Update text and color when a file is selected)
const fileInput = document.getElementById('document');
const fileNameDisplay = document.getElementById('fileNameDisplay');

fileInput.addEventListener('change', function() {
    if(this.files && this.files.length > 0) {
        fileNameDisplay.innerHTML = `<strong><i class="fa-solid fa-file-circle-check"></i> File Attached:</strong> ${this.files[0].name}`;
        fileNameDisplay.style.color = '#00A896'; // Dammam Valley Teal
    } else {
        fileNameDisplay.innerHTML = 'Drag and drop your file here, or click to browse';
        fileNameDisplay.style.color = '#666666';
    }
});

// Power Automate Webhook URL
// IMPORTANT: Keep this URL secure.
const powerAutomateUrl = "https://default58d349d2231941f99afff3055d0a18.80.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/0a66f6833aa04ca68f94ff8d4b84b180/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZO_h_F9xftNmNbcfPkFnrGaqHKIpDdxQ3jWqmIklIkk";

const signatureForm = document.getElementById('signatureForm');
const submitButton = signatureForm.querySelector('.btn-primary');

signatureForm.addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent standard browser form submission

    // UX Enhancement: Show a loading spinner on the submit button
    const originalBtnText = submitButton.innerHTML;
    submitButton.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Sending...';
    submitButton.disabled = true;

    // Retrieve input values
    const recipientInput = document.getElementById('recipient').value;
    const documentFile = document.getElementById('document').files[0];
    const senderEmail = document.getElementById('senderEmail').value;
    const senderName = document.getElementById('senderName').value;

    // Parse the Recipient input (Format expected: "Name (Email)")
    let extractedName = recipientInput;
    let extractedEmail = recipientInput;
    
    if(recipientInput.includes('(') && recipientInput.includes(')')) {
        extractedName = recipientInput.split(' (')[0].trim();
        extractedEmail = recipientInput.substring(
            recipientInput.indexOf("(") + 1, 
            recipientInput.lastIndexOf(")")
        ).trim();
    }

    // Convert the selected file to a Base64 string for API transmission
    const reader = new FileReader();
    reader.readAsDataURL(documentFile);
    
    reader.onload = async function () {
        // Strip out the Data URL metadata prefix to get the raw Base64 string
        const base64FileContent = reader.result.split(',')[1];

        // Construct the JSON payload matching the Power Automate schema
        const payload = {
            senderName: senderName,
            recipientName: extractedName,
            recipientEmail: extractedEmail,
            senderEmail: senderEmail,
            fileName: documentFile.name,
            fileContent: base64FileContent
        };

        try {
            // Execute the POST request to trigger the Flow
            const response = await fetch(powerAutomateUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if(response.ok) {
                alert("Request sent successfully! \nتم إرسال الطلب بنجاح إلى " + extractedName);
                
                // Reset form UI upon success
                signatureForm.reset(); 
                fileNameDisplay.innerHTML = 'Drag and drop your file here, or click to browse';
                fileNameDisplay.style.color = '#666666';
            } else {
                alert("Error sending request. Please check your connection.");
            }
        } catch (error) {
            console.error("Transmission Error:", error);
            alert("An error occurred while sending the document.");
        } finally {
            // Restore the submit button to its original state
            submitButton.innerHTML = originalBtnText;
            submitButton.disabled = false;
        }
    };
});