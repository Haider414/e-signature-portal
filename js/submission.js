/* ==========================================
   7. SUBMISSION MODULE (Modal, Reordering & API POST)
   ========================================== */
const SubmissionModule = (function() {

    // 1. Open options modal and review files & recipients
    function openOptionsModal() {
        const documentFiles = FilesModule.getFiles();
        const recipients = RecipientsModule.getSelectedRecipients();
        
        if (recipients.length === 0) {
            alert("Please select at least one recipient.");
            return;
        }
        if (documentFiles.length === 0) {
            alert("Please attach at least one document.");
            return;
        }

        // Render selected files
        const filesList = document.getElementById('summaryFilesList');
        if (filesList) {
            filesList.innerHTML = '';
            for (let i = 0; i < documentFiles.length; i++) {
                filesList.innerHTML += `<li><i class="fa-solid fa-file-lines" style="color:var(--dv-teal); margin-right:5px;"></i> ${documentFiles[i].name}</li>`;
            }
        }

        // Render recipients with reordering arrows
        renderSummaryRecipients();

        // Show modal
        const modal = document.getElementById('sendOptionsModal');
        if (modal) modal.style.display = 'block';
    }

    // Close options modal
    function closeOptionsModal() {
        const modal = document.getElementById('sendOptionsModal');
        if (modal) modal.style.display = 'none';
    }

    // Render recipients list inside modal with up/down controls
    function renderSummaryRecipients() {
        const list = document.getElementById('summaryRecipientsList');
        if (!list) return;
        list.innerHTML = '';
        
        const recipients = RecipientsModule.getSelectedRecipients();

        recipients.forEach((rec, index) => {
            const li = document.createElement('li');
            
            const info = document.createElement('span');
            info.innerHTML = `<strong>${rec.name}</strong> <span style="font-size:11px; color:#888;">(${rec.email})</span>`;
            
            const controls = document.createElement('div');
            controls.className = 'reorder-controls';
            
            const upBtn = document.createElement('button');
            upBtn.innerHTML = '<i class="fa-solid fa-chevron-up"></i>';
            upBtn.disabled = index === 0;
            upBtn.onclick = () => moveRecipient(index, -1);
            
            const downBtn = document.createElement('button');
            downBtn.innerHTML = '<i class="fa-solid fa-chevron-down"></i>';
            downBtn.disabled = index === recipients.length - 1;
            downBtn.onclick = () => moveRecipient(index, 1);
            
            controls.appendChild(upBtn);
            controls.appendChild(downBtn);
            
            li.appendChild(info);
            li.appendChild(controls);
            list.appendChild(li);
        });
    }

    // Move recipient position up or down
    function moveRecipient(index, direction) {
        const recipients = RecipientsModule.getSelectedRecipients();
        const temp = recipients[index];
        recipients[index] = recipients[index + direction];
        recipients[index + direction] = temp;
        
        renderSummaryRecipients();
    }

    // Final submission to Power Automate webhook
    async function submitFinalRequest(signaturePlacement) {
        const routingInput = document.querySelector('input[name="modalRouting"]:checked');
        const routingType = routingInput ? routingInput.value : 'sequential';
        const documentFiles = FilesModule.getFiles();
        const recipients = RecipientsModule.getSelectedRecipients();
        
        const senderName = document.getElementById('senderName').value || currentUserEmail;
        const senderEmail = document.getElementById('senderEmail').value || currentUserEmail;

        const modalContent = document.querySelector('#sendOptionsModal .modal-content');
        const originalContent = modalContent ? modalContent.innerHTML : '';
        
        if (modalContent) {
            modalContent.innerHTML = '<div style="text-align:center; padding: 40px;"><h2><i class="fa-solid fa-spinner fa-spin" style="color:var(--dv-teal);"></i> Encrypting & Preparing Files...</h2><p>Please wait while we connect to Adobe Sign servers.</p></div>';
        }

        try {
            let filesData = [];
            for (let i = 0; i < documentFiles.length; i++) {
                const base64String = await convertFileToBase64(documentFiles[i]);
                filesData.push({
                    fileName: documentFiles[i].name,
                    fileContent: base64String
                });
            }

            const payload = {
                senderName: senderName,
                senderEmail: senderEmail,
                recipients: recipients,
                routingType: routingType,
                signaturePlacement: signaturePlacement,
                files: filesData
            };

            const response = await fetch(Config.powerAutomateWebhookUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                alert("Request sent successfully! Your workflow has been triggered.");
                
                // Reset form and states
                const form = document.getElementById('signatureForm');
                if (form) form.reset();
                RecipientsModule.clearRecipients();
                FilesModule.clearFiles();
                
                const fileNameDisplay = document.getElementById('fileNameDisplay');
                if (fileNameDisplay) {
                    fileNameDisplay.innerHTML = 'Drag and drop your files here, or click to browse';
                    fileNameDisplay.style.display = 'block';
                }
                
                closeOptionsModal();
            } else {
                throw new Error("Server Response Error");
            }

        } catch (error) {
            console.error("Submission Error:", error);
            alert("An error occurred while sending the request. Please check your connection.");
        } finally {
            if (modalContent) {
                modalContent.innerHTML = originalContent;
            }
        }
    }

    // Helper: Convert file to Base64
    function convertFileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result.split(',')[1]);
            reader.onerror = error => reject(error);
        });
    }

    return {
        openOptionsModal,
        closeOptionsModal,
        submitFinalRequest
    };
})();