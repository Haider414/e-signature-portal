/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    // تهيئة مكتبة MSAL باستخدام الإعدادات من config.js
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // دالة تسجيل الدخول
    async function authenticateUser() {
        try {
            const loginResponse = await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"]
            });
            
            const user = loginResponse.account;
            currentUserEmail = user.username; 
            currentUserName = user.name;       
            
            // تحديث الواجهة (HTML) ببيانات الموظف الحقيقية
            document.getElementById('headerUserName').textContent = currentUserName;
            document.getElementById('menuUserName').textContent = currentUserName;
            document.getElementById('menuUserEmail').textContent = currentUserEmail;
            
            // تحديث الحقول المخفية لنموذج الإرسال
            const senderNameInput = document.getElementById('senderName');
            const senderEmailInput = document.getElementById('senderEmail');
            if (senderNameInput) senderNameInput.value = currentUserName;
            if (senderEmailInput) senderEmailInput.value = currentUserEmail;
            
            console.log("Authentication Successful! User:", currentUserEmail);
            
            // استدعاء دالة تحميل الداشبورد إذا كانت الوحدة (Module) جاهزة
            if (typeof DashboardModule !== 'undefined') {
                DashboardModule.loadDashboardData();
            }
            
        } catch (error) {
            console.error("Authentication Error:", error);
            alert("Authentication failed. Please check Azure AD permissions.");
        }
    }

    // إظهار وإخفاء القائمة المنسدلة للملف الشخصي
    function toggleProfileMenu(event) {
        const menu = document.getElementById('profileMenu');
        if (menu) {
            menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
        }
    }

    // تسجيل الخروج
    function logoutUser() {
        const logoutRequest = {
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        };
        msalInstance.logoutRedirect(logoutRequest);
    }

    // إغلاق القائمة عند النقر في أي مكان خارجها
    document.addEventListener('click', function(event) {
        const profileBtn = document.getElementById('userProfileBtn');
        const profileMenu = document.getElementById('profileMenu');
        
        if (profileMenu && profileBtn && !profileBtn.contains(event.target)) {
            profileMenu.style.display = 'none';
        }
    });

    // تصدير الدوال التي نحتاجها في الملفات الأخرى أو في HTML
    return {
        authenticateUser,
        toggleProfileMenu,
        logoutUser,
        msalInstance // نحتاجه لاحقاً لجلب رمز المصادقة (Token) لـ Graph API
    };
})();