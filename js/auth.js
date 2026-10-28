/* ==========================================
   2. AUTHENTICATION MODULE (Azure AD & MSAL)
   ========================================== */
const AuthModule = (function() {
    
    const msalInstance = new msal.PublicClientApplication(Config.msalConfig);

    // Check initial auth state on load
    function checkUserSession() {
        const accounts = msalInstance.getAllAccounts();
        const authBtn = document.getElementById('authActionButton');
        const profileHeader = document.getElementById('profileMenuHeader');
        const greeting = document.getElementById('profileGreeting');
        const userNameEl = document.getElementById('headerUserName');

        if (accounts.length > 0) {
            const user = accounts[0];
            currentUserEmail = user.username;
            currentUserName = user.name;

            userNameEl.textContent = currentUserName;
            greeting.textContent = "Welcome,";
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
            userNameEl.textContent = "Sign In Required";
            greeting.textContent = "Hello,";
            if (profileHeader) profileHeader.style.display = 'none';

            if (authBtn) {
                authBtn.className = 'btn-signin';
                authBtn.innerHTML = '<i class="fa-solid fa-arrow-right-to-bracket"></i> Sign In';
            }
        }
    }

    // Combined handler for Sign In / Sign Out button
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
                scopes: ["User.Read", "Sites.Read.All", "User.ReadBasic.All"]
            });
            
            checkUserSession();
            console.log("Authentication Successful!");
            
        } catch (error) {
            console.error("Authentication Error:", error);
            alert("Authentication failed. Please check Azure AD permissions.");
        }
    }

    function toggleProfileMenu(event) {
        const menu = document.getElementById('profileMenu');
        if (menu) {
            menu.style.display = menu.style.display === 'none' ? 'block' : 'none';
        }
    }

    function logoutUser() {
        const logoutRequest = {
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        };
        msalInstance.logoutRedirect(logoutRequest).then(() => {
            localStorage.clear();
        }).catch(error => {
            console.error("Logout Error:", error);
        });
    }

    document.addEventListener('click', function(event) {
        const profileBtn = document.getElementById('userProfileBtn');
        const profileMenu = document.getElementById('profileMenu');
        
        if (profileMenu && profileBtn && !profileBtn.contains(event.target)) {
            profileMenu.style.display = 'none';
        }
    });

    // Run check on page load
    window.addEventListener('DOMContentLoaded', () => {
        checkUserSession();
    });

    return {
        authenticateUser,
        toggleProfileMenu,
        logoutUser,
        handleAuthAction,
        msalInstance
    };
})();