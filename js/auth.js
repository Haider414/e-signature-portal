/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام ومحاولة الدخول الصامت (تجهيزاً لبيئة Teams/PowerPages)
    async function initializeAuth() {
        try {
            await msalInstance.handleRedirectPromise();
            
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                // محاولة تجديد الجلسة بصمت في الخلفية إذا كانت صالحة
                await msalInstance.ssoSilent({ account: accounts[0] });
            }
        } catch (error) {
            console.warn("Silent SSO failed or no active session.");
            sessionStorage.clear(); // تنظيف استباقي في حال فشل الجلسة
        } finally {
            checkUserSession();
        }
    }

    // 2. تحديث الواجهة بناءً على حالة الجلسة
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

    // 3. معالج الأزرار
    function handleAuthAction() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            logoutUser();
        } else {
            authenticateUser();
        }
    }

    // 4. تسجيل الدخول النظيف عبر Popup
    async function authenticateUser() {
        try {
            sessionStorage.clear(); // مسح الذاكرة لضمان عدم وجود أقفال وهمية
            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account" // يتيح التبديل بين الحسابات بسهولة
            });
            checkUserSession(); 
        } catch (error) {
            console.error("Popup Login Error:", error);
            const ignoredErrors = ["user_cancelled", "interaction_in_progress", "popup_window_error"];
            // إظهار الخطأ فقط في حال فشل حقيقي وتجاهل الإلغاء اليدوي أو أخطاء التداخل
            if (!ignoredErrors.includes(error.errorCode)) {
                alert("Authentication failed. Please check Azure AD permissions.");
            }
        }
    }

    // 5. تسجيل الخروج الجذري
    async function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) return;

        try {
            await msalInstance.logoutPopup({ account: accounts[0] });
        } catch (error) {
            console.warn("Logout popup failed, cleaning up manually.", error);
        } finally {
            // تدمير الجلسة محلياً بشكل قاطع لضمان نسيان الحساب حتى لو فشلت النافذة
            sessionStorage.clear();
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