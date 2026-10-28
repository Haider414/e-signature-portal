/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. حل مشكلة الجلسات العالقة (Stuck Sessions)
    async function initializeAuth() {
        try {
            await msalInstance.handleRedirectPromise();
            checkUserSession();
        } catch (error) {
            console.error("MSAL Initialization Error:", error);
            checkUserSession();
        }
    }

    // 2. فحص حالة الجلسة وتحديث الواجهة (الاسم، الإيميل، الأنشطة)
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

            // تحديث عناصر الواجهة بمعلومات المستخدم الحقيقية
            if (userNameEl) userNameEl.textContent = currentUserName;
            if (greeting) greeting.textContent = "Welcome,";
            if (menuUserName) menuUserName.textContent = currentUserName;
            if (menuUserEmail) menuUserEmail.textContent = currentUserEmail;
            if (profileHeader) profileHeader.style.display = 'block';

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
            // تصفير الواجهة عند عدم تسجيل الدخول
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

    function handleAuthAction() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            logoutUser();
        } else {
            authenticateUser();
        }
    }

    async function authenticateUser() {
        try {
            const loginResponse = await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account" // إجبار النظام على اختيار الحساب لمنع التداخل
            });
            
            checkUserSession();
            console.log("Authentication Successful!");
            
        } catch (error) {
            console.error("Authentication Error:", error);
            // تجاهل أخطاء التداخل وإغلاق النافذة من قبل المستخدم
            if (error.errorCode !== "user_cancelled" && error.errorCode !== "interaction_in_progress") {
                alert("Authentication failed. Please check Azure AD permissions.");
            }
        }
    }

    function toggleProfileMenu(event) {
        const menu = document.getElementById('profileMenu');
        if (menu) {
            menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
        }
    }

    function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            const logoutRequest = {
                account: accounts[0], // إخبار MSAL بالحساب الدقيق لتسجيل الخروج منه
                postLogoutRedirectUri: window.location.origin + window.location.pathname
            };
            
            // تنظيف الذاكرة المخبأة بالكامل قبل إعادة التوجيه
            sessionStorage.clear();
            localStorage.clear();

            msalInstance.logoutRedirect(logoutRequest);
        }
    }

    document.addEventListener('click', function(event) {
        const profileBtn = document.getElementById('userProfileBtn');
        const profileMenu = document.getElementById('profileMenu');
        
        if (profileMenu && profileBtn && !profileBtn.contains(event.target)) {
            profileMenu.style.display = 'none';
        }
    });

    // استخدام الدالة الجديدة عند تحميل الصفحة
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