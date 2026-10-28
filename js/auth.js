/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام وتنظيف أي بقايا من جلسات مغلقة بشكل مفاجئ
    async function initializeAuth() {
        // تنظيف شريط الرابط من أي أكواد عالقة من مايكروسوفت قبل بدء النظام
        if (window.location.hash.includes("code=") || window.location.hash.includes("state=") || window.location.hash.includes("error=")) {
            window.history.replaceState(null, null, window.location.pathname);
        }

        try {
            await msalInstance.handleRedirectPromise();
        } catch (error) {
            console.warn("MSAL Init Warning:", error);
            // إذا كانت الذاكرة تالفة تماماً، نقوم بمسحها
            sessionStorage.clear();
            localStorage.clear();
        } finally {
            checkUserSession();
        }
    }

    // 2. التحقق من حالة الجلسة وتحديث الواجهة
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

    // 3. معالج النقر على الأزرار
    function handleAuthAction() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            logoutUser();
        } else {
            authenticateUser();
        }
    }

    // 4. الدخول الصافي عبر النافذة المنبثقة (مع كسر الأقفال الوهمية)
    async function authenticateUser() {
        try {
            // كسر القفل الوهمي (interaction_in_progress) الناتج عن إغلاق الصفحة فجأة سابقاً
            sessionStorage.clear();

            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account" 
            });
            checkUserSession(); 
        } catch (error) {
            console.error("Popup Login Error:", error);
            
            // فلتر ذكي لتجاهل الأخطاء الناتجة عن تعليق المتصفح أو إغلاق المستخدم للنافذة
            const ignoredErrors = [
                "user_cancelled", 
                "interaction_in_progress", 
                "popup_window_error",
                "empty_window_error",
                "invalid_state",
                "hash_empty_error"
            ];
            
            if (error.errorCode && !ignoredErrors.includes(error.errorCode)) {
                // في حال وجود خطأ حقيقي، نمسح الذاكرة لكي تنجح المحاولة التالية
                sessionStorage.clear();
                localStorage.clear();
                alert("Authentication failed. Please try again.");
            }
        }
    }

    // 5. الخروج النظيف
    async function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) return;

        try {
            await msalInstance.logoutPopup({
                account: accounts[0]
            });
            sessionStorage.clear();
            localStorage.clear();
            checkUserSession();
        } catch (error) {
            console.warn("Popup blocked or failed, using redirect fallback...", error);
            sessionStorage.clear();
            localStorage.clear();
            msalInstance.logoutRedirect({
                account: accounts[0],
                postLogoutRedirectUri: window.location.origin + window.location.pathname
            });
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