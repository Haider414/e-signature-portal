/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام وتنظيف أي بقايا من عمليات سابقة في الذاكرة
    async function initializeAuth() {
        try {
            await msalInstance.handleRedirectPromise();
        } catch (error) {
            console.warn("MSAL Background Cleanup:", error);
        } finally {
            checkUserSession(); // تحديث الواجهة دائماً سواء نجح التنظيف أم لا
        }
    }

    // 2. التحقق من حالة الجلسة وتحديث الواجهة (الاسم، الايميل، الازرار)
    function checkUserSession() {
        const accounts = msalInstance.getAllAccounts();
        const authBtn = document.getElementById('authActionButton');
        const profileHeader = document.getElementById('profileMenuHeader');
        const greeting = document.getElementById('profileGreeting');
        const userNameEl = document.getElementById('headerUserName');
        const menuUserName = document.getElementById('menuUserName');
        const menuUserEmail = document.getElementById('menuUserEmail');

        if (accounts.length > 0) {
            const user = accounts[0];
            currentUserEmail = user.username;
            currentUserName = user.name;

            if (userNameEl) userNameEl.textContent = currentUserName;
            if (greeting) greeting.textContent = "Welcome,";
            if (menuUserName) menuUserName.textContent = currentUserName;
            if (menuUserEmail) menuUserEmail.textContent = currentUserEmail;
            if (profileHeader) profileHeader.style.display = 'flex'; 

            if (authBtn) {
                authBtn.className = 'btn-logout';
                authBtn.innerHTML = '<i class="fa-solid fa-arrow-right-from-bracket"></i> Sign Out';
            }

            const senderNameInput = document.getElementById('senderName');
            const senderEmailInput = document.getElementById('senderEmail');
            if (senderNameInput) senderNameInput.value = currentUserName;
            if (senderEmailInput) senderEmailInput.value = currentUserEmail;

            if (typeof DashboardModule !== 'undefined') {
                DashboardModule.loadDashboardData();
            }
        } else {
            if (userNameEl) userNameEl.textContent = "Sign In Required";
            if (greeting) greeting.textContent = "Hello,";
            
            if (menuUserName) menuUserName.textContent = '';
            if (menuUserEmail) menuUserEmail.textContent = '';
            if (profileHeader) profileHeader.style.display = 'none';

            if (authBtn) {
                authBtn.className = 'btn-signin';
                authBtn.innerHTML = '<i class="fa-solid fa-arrow-right-to-bracket"></i> Sign In';
            }
            
            const recentContainer = document.getElementById('recentActivityContainer');
            if(recentContainer) {
                 recentContainer.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">Sign in to view recent activity.</p>';
            }
        }
    }

    // 3. معالج النقر على زر (تسجيل الدخول / الخروج)
    function handleAuthAction() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            logoutUser();
        } else {
            authenticateUser();
        }
    }

    // 4. تسجيل الدخول عبر النافذة المنبثقة (Popup) بدون حلقات تكرار لا نهائية
    async function authenticateUser() {
        try {
            // تنظيف الذاكرة المؤقتة لمنع خطأ interaction_in_progress قبل فتح النافذة
            sessionStorage.clear();
            
            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account" // يتيح لك اختيار الحساب دائماً
            });
            
            // تحديث الواجهة فوراً بعد نجاح الدخول
            checkUserSession();
            
        } catch (error) {
            console.error("Popup Login Action:", error.errorCode);
            
            // الفلتر الذكي: نتجاهل جميع الأخطاء الوهمية والمزعجة بصمت تام
            const ignoredErrors = [
                "user_cancelled", 
                "interaction_in_progress", 
                "popup_window_error", 
                "empty_window_error",
                "monitor_window_timeout"
            ];
            
            if (!ignoredErrors.includes(error.errorCode)) {
                console.warn("Authentication interrupted:", error);
            }
        }
    }

    // 5. تسجيل الخروج السليم
    function logoutUser() {
        try {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                const logoutRequest = {
                    account: accounts[0],
                    postLogoutRedirectUri: window.location.origin + window.location.pathname
                };
                
                // تنظيف الذاكرة بالكامل لضمان نسيان الحساب
                sessionStorage.clear();
                localStorage.clear();

                msalInstance.logoutRedirect(logoutRequest);
            }
        } catch (error) {
            console.error("Logout execution error:", error);
            // في حال حدوث خطأ، نقوم بتحديث الواجهة للخروج محلياً على الأقل
            sessionStorage.clear();
            localStorage.clear();
            checkUserSession();
        }
    }

    // 6. التحكم بالقائمة المنسدلة
    function toggleProfileMenu(event) {
        const menu = document.getElementById('profileMenu');
        if (menu) {
            menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
        }
    }

    document.addEventListener('click', function(event) {
        const profileBtn = document.getElementById('userProfileBtn');
        const profileMenu = document.getElementById('profileMenu');
        
        if (profileMenu && profileBtn && !profileBtn.contains(event.target)) {
            profileMenu.style.display = 'none';
        }
    });

    // تشغيل التهيئة عند تحميل الصفحة
    window.addEventListener('DOMContentLoaded', () => {
        initializeAuth();
    });

    return {
        authenticateUser,
        toggleProfileMenu,
        logoutUser,
        handleAuthAction,
        msalInstance
    };
})();