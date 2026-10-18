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
        knownAuthorities: ["login.microsoftonline.com"], // هذا السطر يحل مشكلة الثقة في المتصفحات
        redirectUri: "https://haider414.github.io/e-signature-portal/" 
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let currentUserEmail = ""; // سنملأ هذا المتغير برمجياً بعد تسجيل الدخول
// مصفوفة لحفظ البيانات الحقيقية القادمة من SharePoint
let realRequests = [];

// 2. دالة تسجيل الدخول التلقائي
async function authenticateUser() {
    try {
        // طلب تسجيل الدخول وقراءة البيانات
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"]
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

/**
 * ==========================================
 * 3. DASHBOARD RENDERING (Fetching from SharePoint)
 * ==========================================
 */

async function loadDashboardData() {
    const sentBody = document.getElementById('sentRequestsBody');
    const receivedBody = document.getElementById('receivedRequestsBody');
    
    sentBody.innerHTML = '<tr><td colspan="5">جاري التحميل من SharePoint...</td></tr>'; 
    receivedBody.innerHTML = '<tr><td colspan="5">جاري التحميل من SharePoint...</td></tr>'; 

    try {
        const account = msalInstance.getAllAccounts()[0];
        if (!account) return;

        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["Sites.Read.All"],
            account: account
        });

        const siteHostname = "dammamvalleyd.sharepoint.com"; 
        const sitePath = "/sites/E-SignatureSystem";
        const listId = "4de9b715-9c83-4c02-a4aa-670346ce2254"; 
        
        const graphUrl = `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}:/lists/${listId}/items?$expand=fields`;
        
        const response = await fetch(graphUrl, {
            headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
        });

        if (!response.ok) throw new Error("خطأ في جلب البيانات");

        const data = await response.json();
        const items = data.value;

        sentBody.innerHTML = '';
        receivedBody.innerHTML = '';
        
        // تفريغ المصفوفة قبل تعبئتها من جديد
        realRequests = []; 

        if(items.length === 0){
             sentBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">لا توجد طلبات مرسلة</td></tr>';
             receivedBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">لا توجد طلبات مستلمة</td></tr>';
             return;
        }

        items.forEach((item, index) => {
            const fields = item.fields;
            const request = {
                id: fields.OrderNumber || fields.AgreementID || fields.OData__x004f_rderNumber || "N/A", 
                senderName: fields.SenderName || fields.Sender_x0020_Name || "N/A",
                senderEmail: fields.SenderEmail || fields.Sender_x0020_Email || "N/A",
                recipientName: fields.RecipientName || fields.Recipient_x0020_Name || fields.RecipientEmail || "N/A",
                recipientEmail: fields.RecipientEmail || fields.Recipient_x0020_Email || "N/A",
                fileName: fields.Title || "N/A", 
                date: new Date(item.createdDateTime).toLocaleDateString(),
                status: fields.RequestStatus || fields.Request_x0020_Status || "Pending" 
            };

            // حفظ الطلب في المصفوفة الحقيقية
            realRequests.push(request);
            // تحديد رقم فهرس هذا الطلب في المصفوفة
            const currentIndex = realRequests.length - 1;

            let statusClass = 'pending';
            if (request.status === 'Completed' || request.status === 'Signed') statusClass = 'signed';
            if (request.status === 'Declined') statusClass = 'declined';
            
            let tr = document.createElement('tr');
            
            if (request.senderEmail === currentUserEmail) {
                tr.innerHTML = `
                    <td>${request.recipientName}</td>
                    <td>${request.fileName}</td>
                    <td>${request.date}</td>
                    <td><span class="status ${statusClass}">${request.status}</span></td>
                    <td><button class="btn-view" onclick="openDynamicModal(${currentIndex})">View Details</button></td>
                `;
                sentBody.appendChild(tr);
            } 
            else if (request.recipientEmail === currentUserEmail) {
                tr.innerHTML = `
                    <td>${request.senderName}</td>
                    <td>${request.fileName}</td>
                    <td>${request.date}</td>
                    <td><span class="status ${statusClass}">${request.status}</span></td>
                    <td><button class="btn-view" onclick="openDynamicModal(${currentIndex})">View Details</button></td>
                `;
                receivedBody.appendChild(tr);
            }
        });
        
    } catch (error) {
        console.error(error);
        sentBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">حدث خطأ أثناء الاتصال بقاعدة البيانات.</td></tr>';
    }
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
    // التعديل هنا: قراءة البيانات من المصفوفة الحقيقية
    const request = realRequests[index];
    
    document.getElementById('modalSenderName').textContent = request.senderName;
    document.getElementById('modalSenderEmail').textContent = request.senderEmail;
    
    document.getElementById('modalRecipientName').textContent = request.recipientName;
    document.getElementById('modalRecipientEmail').textContent = request.recipientEmail;
    
    document.getElementById('modalFileName').textContent = request.fileName;
    document.getElementById('modalOrderId').textContent = request.id;
    document.getElementById('modalDate').textContent = request.date;
    
    const statusBadge = document.getElementById('modalStatusBadge');
    statusBadge.textContent = request.status;
    
    // دعم حالة Completed
    if (request.status === 'Signed' || request.status === 'Completed') {
        statusBadge.className = 'status signed';
    } else if (request.status === 'Declined') {
        statusBadge.className = 'status declined'; 
    } else {
        statusBadge.className = 'status pending';
    }
    
    const btnReminder = document.getElementById('btnSendReminder');
    if (request.status === 'Pending' && request.senderEmail === currentUserEmail) {
        btnReminder.style.display = 'block'; 
    } else {
        btnReminder.style.display = 'none';  
    }

    const btnView = document.getElementById('btnViewDocument');
    btnView.onclick = function() {
        alert("This feature will be connected to SharePoint soon to view: " + request.fileName);
    };
    
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

/**
 * ==========================================
 * 9. SMART SEARCH & CHIPS MANAGEMENT
 * ==========================================
 */
const searchInput = document.getElementById('recipientSearch');
const dropdown = document.getElementById('searchResultsDropdown');
const chipsContainer = document.getElementById('recipientChipsContainer');
const finalRecipients = document.getElementById('finalRecipients');

let selectedRecipients = []; // مصفوفة لتخزين المستلمين المختارين

// استدعاء Graph API للبحث عن موظف
async function searchUsers(searchTerm) {
    if(searchTerm.length < 2) {
        dropdown.style.display = 'none';
        return;
    }
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadBasic.All"], account: account
        });
        
        // جلب الموظفين الذين تبدأ أسماؤهم أو إيميلاتهم بنص البحث
        const url = `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${searchTerm}') or startswith(mail,'${searchTerm}')&$top=5`;
        const response = await fetch(url, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
        const data = await response.json();
        
        renderDropdown(data.value);
    } catch (error) {
        console.error("Search Error:", error);
    }
}

// عرض النتائج في القائمة المنسدلة
function renderDropdown(users) {
    dropdown.innerHTML = '';
    if(!users || users.length === 0) {
        dropdown.style.display = 'none';
        return;
    }
    
    users.forEach(user => {
        let div = document.createElement('div');
        div.className = 'dropdown-item';
        div.innerHTML = `<span class="dropdown-item-name">${user.displayName}</span><span class="dropdown-item-email">${user.mail}</span>`;
        
        div.onclick = function() {
            addChip(user.displayName, user.mail);
            dropdown.style.display = 'none';
            searchInput.value = '';
        };
        dropdown.appendChild(div);
    });
    dropdown.style.display = 'block';
}

// إضافة Chip
function addChip(name, email) {
    // منع التكرار
    if(selectedRecipients.some(r => r.email === email)) return;
    
    selectedRecipients.push({ name: name, email: email });
    updateChipsUI();
}

// إزالة Chip
function removeChip(email) {
    selectedRecipients = selectedRecipients.filter(r => r.email !== email);
    updateChipsUI();
}

// تحديث واجهة الـ Chips وقيمة الحقل المخفي
function updateChipsUI() {
    // مسح الـ Chips القديمة (مع الحفاظ على مربع البحث)
    document.querySelectorAll('.chip').forEach(el => el.remove());
    
    selectedRecipients.forEach(user => {
        let chip = document.createElement('div');
        chip.className = 'chip';
        chip.innerHTML = `<i class="fa-solid fa-user"></i> ${user.name} <span class="chip-close" onclick="removeChip('${user.email}')">&times;</span>`;
        chipsContainer.insertBefore(chip, searchInput);
    });
    
    // تحويل المصفوفة إلى نص (JSON) لإرسالها لـ Power Automate
    finalRecipients.value = JSON.stringify(selectedRecipients);
}

// مراقبة الكتابة في مربع البحث (مع تأخير بسيط Debounce لعدم إرهاق السيرفر)
let searchTimeout;
searchInput.addEventListener('input', function() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => searchUsers(this.value), 400);
});

// إغلاق القائمة المنسدلة عند النقر في أي مكان آخر
document.addEventListener('click', function(e) {
    if(e.target !== searchInput && e.target !== dropdown) {
        dropdown.style.display = 'none';
    }
});