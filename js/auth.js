/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام
    async function initializeAuth() {
        try {
            await msalInstance.handleRedirectPromise();
        } catch (error) {
            console.warn("MSAL Init:", error);
        } finally {
            checkUserSession();
        }
    }

    // 2. تحديث الواجهة بناءً على حالة الجلسة الحالية
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

    // 4. تسجيل الدخول الذكي (مع كاسر الأقفال الوهمية)
    async function authenticateUser() {
        try {
            // التنظيف الاستباقي: إزالة قفل MSAL الوهمي الذي يمنع الدخول المباشر بعد الخروج
            sessionStorage.removeItem("msal.interaction.status");

            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account"
            });
            checkUserSession(); 
        } catch (error) {
            // معالجة الأخطاء الذكية
            if (error.errorCode === "interaction_in_progress") {
                // إذا أصر النظام على وجود قفل، نمسح الذاكرة بالكامل لإنقاذه
                sessionStorage.clear();
            } else if (error.errorCode !== "user_cancelled") {
                alert("Authentication failed. Please try again.");
            }
            console.warn("Login Action:", error);
        }
    }

    // 5. تسجيل الخروج الجذري
    async function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) return;

        try {
            await msalInstance.logoutPopup({
                account: accounts[0]
            });
        } catch (error) {
            console.warn("Logout popup:", error);
        } finally {
            // السر هنا: تنظيف الذاكرة المؤقتة *بعد* انتهاء عملية الخروج بالكامل
            // هذا يضمن أن النظام جاهز 100% لاستقبال تسجيل دخول جديد وفوري
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