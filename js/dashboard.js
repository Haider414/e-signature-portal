/* ==========================================
   4. DASHBOARD MODULE (SharePoint Integration & Tables)
   ========================================== */
const DashboardModule = (function() {

    // Fetch and render sent/received requests from SharePoint list
    async function loadDashboardData() {
        const sentBody = document.getElementById('sentRequestsBody');
        const receivedBody = document.getElementById('receivedRequestsBody');
        
        if (!sentBody || !receivedBody) return;

        sentBody.innerHTML = '<tr><td colspan="5">Loading from SharePoint...</td></tr>'; 
        receivedBody.innerHTML = '<tr><td colspan="5">Loading from SharePoint...</td></tr>'; 

        try {
            const account = AuthModule.msalInstance.getAllAccounts()[0];
            if (!account) return;

            const tokenResponse = await AuthModule.msalInstance.acquireTokenSilent({
                scopes: ["Sites.Read.All"],
                account: account
            });

            const { siteHostname, sitePath, listId } = Config.sharePointContext;
            const graphUrl = `${Config.graphApiUrl}/sites/${siteHostname}:${sitePath}:/lists/${listId}/items?$expand=fields`;
            
            const response = await fetch(graphUrl, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });

            if (!response.ok) throw new Error("Failed to fetch data from SharePoint");

            const data = await response.json();
            const items = data.value;

            sentBody.innerHTML = '';
            receivedBody.innerHTML = '';
            
            realRequests = []; // Clear array before repopulating

            if (items.length === 0) {
                 sentBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No sent requests found</td></tr>';
                 receivedBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No received requests found</td></tr>';
                 return;
            }

            items.forEach((item) => {
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

                realRequests.push(request);
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
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})">View Details</button></td>
                    `;
                    sentBody.appendChild(tr);
                } 
                else if (request.recipientEmail === currentUserEmail) {
                    tr.innerHTML = `
                        <td>${request.senderName}</td>
                        <td>${request.fileName}</td>
                        <td>${request.date}</td>
                        <td><span class="status ${statusClass}">${request.status}</span></td>
                        <td><button class="btn-view" onclick="DashboardModule.openDynamicModal(${currentIndex})">View Details</button></td>
                    `;
                    receivedBody.appendChild(tr);
                }
            });
            
        } catch (error) {
            console.error("Dashboard Error:", error);
            sentBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
            receivedBody.innerHTML = '<tr><td colspan="5" style="color:red; text-align:center;">Error connecting to database.</td></tr>';
        }
    }

    // Open details modal for a specific request
    function openDynamicModal(index) {
        const request = realRequests[index];
        if (!request) return;
        
        document.getElementById('modalSenderName').textContent = request.senderName;
        document.getElementById('modalSenderEmail').textContent = request.senderEmail;
        
        document.getElementById('modalRecipientName').textContent = request.recipientName;
        document.getElementById('modalRecipientEmail').textContent = request.recipientEmail;
        
        document.getElementById('modalFileName').textContent = request.fileName;
        document.getElementById('modalOrderId').textContent = request.id;
        document.getElementById('modalDate').textContent = request.date;
        
        const statusBadge = document.getElementById('modalStatusBadge');
        statusBadge.textContent = request.status;
        
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
            alert("Viewing document: " + request.fileName);
        };
        
        document.getElementById('detailsModal').style.display = 'block';
    }

    // Close details modal
    function closeModal() {
        const modal = document.getElementById('detailsModal');
        if (modal) modal.style.display = 'none';
    }

    // Close modal when clicking outside of content
    window.addEventListener('click', function(event) {
        const modal = document.getElementById('detailsModal');
        if (event.target === modal) {
            modal.style.display = "none";
        }
    });

    return {
        loadDashboardData,
        openDynamicModal,
        closeModal
    };
})();