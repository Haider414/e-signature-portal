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
        document.getElementById('menuUserName').textContent = fullName;
        document.getElementById('menuUserEmail').textContent = currentUserEmail;
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
 * 11. USER PROFILE & LOGOUT MANAGEMENT
 * ==========================================
 */

// إظهار وإخفاء القائمة المنسدلة
function toggleProfileMenu(event) {
    const menu = document.getElementById('profileMenu');
    menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
}

// إغلاق القائمة عند النقر في أي مكان خارجها
document.addEventListener('click', function(event) {
    const profileBtn = document.getElementById('userProfileBtn');
    const profileMenu = document.getElementById('profileMenu');
    
    if (profileMenu && !profileBtn.contains(event.target)) {
        profileMenu.style.display = 'none';
    }
});

// دالة تسجيل الخروج من تطبيق Azure (MSAL)
function logoutUser() {
    // استخدم رابط موقعك الحالي (GitHub Pages) للعودة إليه بعد تسجيل الخروج
    const logoutRequest = {
        postLogoutRedirectUri: window.location.origin + window.location.pathname
    };
    msalInstance.logoutRedirect(logoutRequest);
}

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
 * 5. FILE UPLOAD & MANAGEMENT SYSTEM
 * ==========================================
 */
const fileInput = document.getElementById('document');
const fileNameDisplay = document.getElementById('fileNameDisplay');
const filesContainer = document.getElementById('uploadedFilesContainer');

// DataTransfer object لحفظ وإدارة الملفات برمجياً 
// (يسمح بالإضافة والحذف المتكرر دون فقدان الملفات السابقة)
let dt = new DataTransfer(); 

fileInput.addEventListener('change', function() {
    // إضافة الملفات الجديدة إلى الكائن dt
    for (let i = 0; i < this.files.length; i++) {
        // يمكن إضافة تحقق من نوع الملف (PDF, Word) أو حجمه هنا
        dt.items.add(this.files[i]);
    }
    
    // تحديث ملفات الحقل الفعلي (input) لتطابق ملفات الـ dt
    this.files = dt.files;
    
    // إعادة رسم قائمة الملفات
    renderFilesList();
});

// دالة لرسم قائمة الملفات في الواجهة
function renderFilesList() {
    filesContainer.innerHTML = ''; // مسح القائمة الحالية
    
    if (dt.files.length > 0) {
        fileNameDisplay.style.display = 'none'; // إخفاء نص "اسحب وأفلت"
        
        for (let i = 0; i < dt.files.length; i++) {
            const file = dt.files[i];
            
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            
            // أيقونة تتغير حسب نوع الملف (اختياري)
            let iconClass = 'fa-file-lines';
            if (file.type.includes('pdf')) iconClass = 'fa-file-pdf';
            if (file.type.includes('word') || file.name.endsWith('.docx')) iconClass = 'fa-file-word';
            
            fileItem.innerHTML = `
                <div class="file-item-info">
                    <i class="fa-solid ${iconClass}"></i> 
                    <span>${file.name} <small style="color:#888;">(${(file.size / 1024).toFixed(1)} KB)</small></span>
                </div>
                <i class="fa-solid fa-xmark file-item-remove" onclick="removeFile(${i}, event)" title="Remove File"></i>
            `;
            
            filesContainer.appendChild(fileItem);
        }
    } else {
        fileNameDisplay.style.display = 'block'; // إظهار النص إذا كانت القائمة فارغة
    }
}

// دالة لحذف ملف معين
function removeFile(index, event) {
    event.stopPropagation(); // منع فتح نافذة اختيار الملفات عند الضغط على زر X
    
    // إزالة الملف من كائن الـ DataTransfer
    dt.items.remove(index);
    
    // تحديث الحقل الفعلي
    fileInput.files = dt.files;
    
    // إعادة رسم القائمة
    renderFilesList();
}

// Power Automate Webhook URL
// IMPORTANT: Keep this URL secure.
const powerAutomateUrl = "https://default58d349d2231941f99afff3055d0a18.80.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/0a66f6833aa04ca68f94ff8d4b84b180/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZO_h_F9xftNmNbcfPkFnrGaqHKIpDdxQ3jWqmIklIkk";

const signatureForm = document.getElementById('signatureForm');
const submitButton = signatureForm.querySelector('.btn-primary');

/**
 * ==========================================
 * 10. SENDING OPTIONS MODAL & FINAL SUBMIT
 * ==========================================
 */

// 1. فتح نافذة الخيارات وتعبئة البيانات (Review & Configure)
function openOptionsModal() {
    const documentFiles = document.getElementById('document').files;
    
    // التحقق من وجود مستلمين وملفات قبل المتابعة
    if (selectedRecipients.length === 0) {
        alert("الرجاء اختيار مستلم واحد على الأقل.");
        return;
    }
    if (documentFiles.length === 0) {
        alert("الرجاء إرفاق ملف واحد على الأقل.");
        return;
    }

    // عرض الملفات في النافذة
    const filesList = document.getElementById('summaryFilesList');
    filesList.innerHTML = '';
    for (let i = 0; i < documentFiles.length; i++) {
        filesList.innerHTML += `<li><i class="fa-solid fa-file-lines" style="color:var(--dv-teal); margin-right:5px;"></i> ${documentFiles[i].name}</li>`;
    }

    // عرض الأسماء مع أزرار الترتيب
    renderSummaryRecipients();

    // إظهار النافذة
    document.getElementById('sendOptionsModal').style.display = 'block';
}

function closeOptionsModal() {
    document.getElementById('sendOptionsModal').style.display = 'none';
}

// 2. دالة رسم قائمة المستلمين مع أسهم الترتيب
function renderSummaryRecipients() {
    const list = document.getElementById('summaryRecipientsList');
    list.innerHTML = '';
    
    selectedRecipients.forEach((rec, index) => {
        const li = document.createElement('li');
        
        // تفاصيل الموظف
        const info = document.createElement('span');
        info.innerHTML = `<strong>${rec.name}</strong> <span style="font-size:11px; color:#888;">(${rec.email})</span>`;
        
        // أزرار التحكم (أسهم لأعلى ولأسفل)
        const controls = document.createElement('div');
        controls.className = 'reorder-controls';
        
        const upBtn = document.createElement('button');
        upBtn.innerHTML = '<i class="fa-solid fa-chevron-up"></i>';
        upBtn.disabled = index === 0; // تعطيل السهم الأعلى لأول شخص
        upBtn.onclick = () => moveRecipient(index, -1);
        
        const downBtn = document.createElement('button');
        downBtn.innerHTML = '<i class="fa-solid fa-chevron-down"></i>';
        downBtn.disabled = index === selectedRecipients.length - 1; // تعطيل السهم الأسفل لآخر شخص
        downBtn.onclick = () => moveRecipient(index, 1);
        
        controls.appendChild(upBtn);
        controls.appendChild(downBtn);
        
        li.appendChild(info);
        li.appendChild(controls);
        list.appendChild(li);
    });
}

// 3. دالة تحريك الموظف في الترتيب (لأعلى أو لأسفل)
function moveRecipient(index, direction) {
    // تبديل المواقع في المصفوفة
    const temp = selectedRecipients[index];
    selectedRecipients[index] = selectedRecipients[index + direction];
    selectedRecipients[index + direction] = temp;
    
    // إعادة رسم القائمة فوراً
    renderSummaryRecipients();
}

// 4. الإرسال النهائي إلى Power Automate
async function submitFinalRequest(signaturePlacement) {
    const routingType = document.querySelector('input[name="modalRouting"]:checked').value;
    const documentFiles = document.getElementById('document').files;
    
    const senderName = document.getElementById('senderName').value || currentUserEmail;
    const senderEmail = document.getElementById('senderEmail').value || currentUserEmail;

    // تغيير شكل الواجهة أثناء التحميل
    const modalContent = document.querySelector('#sendOptionsModal .modal-content');
    const originalContent = modalContent.innerHTML;
    modalContent.innerHTML = '<div style="text-align:center; padding: 40px;"><h2><i class="fa-solid fa-spinner fa-spin" style="color:var(--dv-teal);"></i> جاري إعداد وتشفير الملفات...</h2><p>يرجى الانتظار، يتم الآن الاتصال بخوادم Adobe Sign</p></div>';

    try {
        // تحويل جميع الملفات المرفقة إلى Base64
        let filesData = [];
        for (let i = 0; i < documentFiles.length; i++) {
            const base64String = await convertFileToBase64(documentFiles[i]);
            filesData.push({
                fileName: documentFiles[i].name,
                fileContent: base64String
            });
        }

        // تجهيز حزمة البيانات الكاملة
        const payload = {
            senderName: senderName,
            senderEmail: senderEmail,
            recipients: selectedRecipients, // مصفوفة الأسماء بالترتيب الجديد
            routingType: routingType,       // sequential أو parallel
            signaturePlacement: signaturePlacement, // auto أو manual
            files: filesData
        };

        // إرسال البيانات
        const response = await fetch(powerAutomateUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if(response.ok) {
            alert("تم إرسال الطلب بنجاح! سيتم توجيهه حسب إعداداتك.");
            // إعادة ضبط كل شيء
            document.getElementById('signatureForm').reset();
            selectedRecipients = [];
            updateChipsUI();
            document.getElementById('fileNameDisplay').innerHTML = 'Drag and drop your files here, or click to browse';
            closeOptionsModal();
        } else {
            throw new Error("Server Response Error");
        }

    } catch (error) {
        console.error("Submission Error:", error);
        alert("حدث خطأ أثناء الإرسال. يرجى التأكد من اتصالك أو التحقق من Power Automate.");
    } finally {
        // إعادة النافذة لحالتها الطبيعية في حال حدث خطأ
        modalContent.innerHTML = originalContent;
    }
}

// دالة مساعدة لتحويل الملفات
function convertFileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    });
}

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
// استدعاء Graph API للبحث الذكي عن موظف
async function searchUsers(searchTerm) {
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadBasic.All"], account: account
        });
        
        // استخدام البحث المتقدم ($search) مع دعم ($count) لجلب أوسع نطاق من النتائج
        let url = `https://graph.microsoft.com/v1.0/users?$top=20`;
        
        // إذا كان هناك نص، نطلب من السيرفر البحث المتقدم
        if (searchTerm.length > 0) {
            url += `&$search="displayName:${searchTerm}" OR "mail:${searchTerm}"&$count=true`;
        }

        const response = await fetch(url, { 
            headers: { 
                "Authorization": `Bearer ${tokenResponse.accessToken}`,
                "ConsistencyLevel": "eventual" // هذا الرأس ضروري جداً لعمل البحث المتقدم في مايكروسوفت
            } 
        });
        const data = await response.json();
        
        // فلترة إضافية فائقة الدقة في الجافاسكريبت (Fuzzy Search)
        // لضمان التقاط الكلمة سواء كانت في بداية الاسم، وسطه، أو نهايته
        let finalUsers = data.value;
        if (searchTerm.length > 0) {
            const term = searchTerm.toLowerCase();
            finalUsers = finalUsers.filter(u => 
                (u.displayName && u.displayName.toLowerCase().includes(term)) || 
                (u.mail && u.mail.toLowerCase().includes(term))
            );
        }

        renderDropdown(finalUsers);
    } catch (error) {
        console.error("Search Error:", error);
    }
}

// عرض الاقتراحات فوراً بمجرد النقر على الحقل (بدون الحاجة للكتابة)
searchInput.addEventListener('focus', function() {
    searchUsers(this.value);
});

// تقليل وقت التأخير (Debounce) ليكون البحث سريعاً جداً أثناء الكتابة
let searchTimeout;
searchInput.addEventListener('input', function() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => searchUsers(this.value), 200); // 200 ملي ثانية فقط
});

// مستمع لحدث الضغط على Enter في مربع البحث
searchInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault(); // منع إرسال النموذج (Form Submit)
        const val = this.value.trim();
        
        // التحقق البسيط من صيغة الإيميل
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        
        if (emailRegex.test(val)) {
            // إذا كان النص إيميلاً صالحاً، نضيفه كـ Chip (الاسم والإيميل متطابقان)
            addChip(val, val);
            this.value = ''; // تفريغ المربع
            dropdown.style.display = 'none'; // إخفاء القائمة المنسدلة
        } else if (val.length > 0) {
            alert("يرجى إدخال بريد إلكتروني صحيح (مثل: name@company.com) لإضافته يدوياً.");
        }
    }
});

// إغلاق القائمة المنسدلة عند النقر في أي مكان آخر
document.addEventListener('click', function(e) {
    if(e.target !== searchInput && e.target !== dropdown) {
        dropdown.style.display = 'none';
    }
});

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


// إغلاق القائمة المنسدلة عند النقر في أي مكان آخر
document.addEventListener('click', function(e) {
    if(e.target !== searchInput && e.target !== dropdown) {
        dropdown.style.display = 'none';
    }
});

