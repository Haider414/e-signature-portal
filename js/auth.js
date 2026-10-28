/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    // تهيئة مكتبة مايكروسوفت بناءً على الإعدادات الجديدة في config.js (والتي تستخدم localStorage)
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // 1. تهيئة النظام ومعالجة أي توجيهات سابقة
    async function initializeAuth() {
        try {
            await msalInstance.handleRedirectPromise();
        } catch (error) {
            console.warn("MSAL Init Warning:", error);
        } finally {
            checkUserSession();
        }
    }

    // 2. التحقق من الجلسة (تعمل بذكاء مع localStorage للاحتفاظ بالدخول بعد إغلاق الصفحة)
    function checkUserSession() {
        const accounts = msalInstance.getAllAccounts();
        const authBtn = document.getElementById('authActionButton');
        const profileHeader = document.getElementById('profileMenuHeader');
        const greeting = document.getElementById('profileGreeting');
        const userNameEl = document.getElementById('headerUserName');
        const menuUserName = document.getElementById('menuUserName');
        const menuUserEmail = document.getElementById('menuUserEmail');

        if (accounts.length > 0) {
            // المستخدم مسجل دخوله (سواء للتو، أو من جلسة سابقة لم يخرج منها)
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

            // تحميل بيانات الداشبورد لأن المستخدم مصرح له
            if (typeof DashboardModule !== 'undefined') {
                DashboardModule.loadDashboardData();
            }
        } else {
            // المستخدم غير مسجل دخوله (أو قام بتسجيل الخروج مسبقاً)
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

    // 4. الدخول الصافي والذكي عبر النافذة المنبثقة
    async function authenticateUser() {
        try {
            await msalInstance.loginPopup({
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"],
                prompt: "select_account" // السر هنا: إجبار النظام على سؤالك أي حساب تريد استخدامه لتسهيل تبديل الحسابات
            });
            checkUserSession(); 
        } catch (error) {
            console.error("Popup Login Error:", error);
            const ignoredErrors = ["user_cancelled", "interaction_in_progress", "popup_window_error"];
            // تجاهل الأخطاء الناتجة عن إغلاق المستخدم للنافذة أو بطء المتصفح بصمت
            if (!ignoredErrors.includes(error.errorCode)) {
                alert("Authentication failed. Please check your connection or permissions.");
            }
        }
    }

    // 5. الخروج النظيف ذو الخطة البديلة (Fallback)
    async function logoutUser() {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) return;

        try {
            // المحاولة الأولى: خروج سلس وسريع عبر النافذة المنبثقة
            await msalInstance.logoutPopup({
                account: accounts[0]
            });
            // مسح أي بقايا في الذاكرة المؤقتة تأكيداً لنظافة الجلسة
            sessionStorage.clear();
            checkUserSession();
        } catch (error) {
            console.warn("Popup blocked or failed, using redirect fallback...", error);
            // المحاولة الثانية (ذكاء الكود): إذا حظر المتصفح النافذة، نستخدم التوجيه الإجباري
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

    // تشغيل النظام فور تحميل الصفحة
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