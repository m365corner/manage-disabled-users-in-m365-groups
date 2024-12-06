const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<client-id-goes-here>",
        authority: "https://login.microsoftonline.com/<tenant-id-goes-here>",
        redirectUri: "http://localhost:8000",
    },
});

let allGroups = [];
let disabledUsers = [];

// Login and Fetch Data
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["Group.Read.All", "User.Read.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
        await fetchGroups();
        await fetchDisabledUsers();
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}

function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch Groups
async function fetchGroups() {
    const response = await callGraphApi("/groups?$select=id,displayName,mail,description,mailEnabled,resourceProvisioningOptions");
    allGroups = response.value.filter(group => !group.resourceProvisioningOptions.includes("Team"));
    populateDropdown("groupDropdown", allGroups.map(group => ({ id: group.id, name: group.displayName })));
}

// Fetch Disabled Users
async function fetchDisabledUsers() {
    const response = await callGraphApi("/users?$filter=accountEnabled eq false&$select=id,displayName,mail");
    disabledUsers = response.value;
    populateDropdown("memberDropdown", disabledUsers.map(user => ({ id: user.id, name: user.displayName })));
}

// Populate Dropdown
function populateDropdown(dropdownId, items) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = `<option value="">Select</option>`;
    items.forEach(item => {
        const option = document.createElement("option");
        option.value = item.id;
        option.textContent = item.name;
        dropdown.appendChild(option);
    });
}

// Search Function


async function search() {
    const groupId = document.getElementById("groupDropdown").value;
    const userId = document.getElementById("memberDropdown").value;
    const mailEnabled = document.getElementById("mailEnabledDropdown").value;

    if (!groupId && !userId && !mailEnabled) {
        alert("Please select at least one filter.");
        return;
    }

    const results = [];

    // If all three filters (Group, Member, Mail Enabled) are selected
    if (groupId && userId && mailEnabled) {
        const selectedGroup = allGroups.find(group => group.id === groupId);
        const selectedUser = disabledUsers.find(user => user.id === userId);

        if (!selectedGroup || !selectedUser) {
            alert("Invalid group or member selection.");
            return;
        }

        if (selectedGroup.mailEnabled.toString() !== mailEnabled) {
            alert("The selected group does not match the mail-enabled filter.");
            return;
        }

        // Fetch members of the selected group
        const members = await callGraphApi(`/groups/${groupId}/members?$select=id,displayName,mail,accountEnabled`);
        const matchingMember = members.value.find(member =>
            member.id === userId && member["@odata.type"] === "#microsoft.graph.user" && member.accountEnabled === false
        );

        if (!matchingMember) {
            alert("The selected user does not belong to the selected group or does not meet the criteria.");
            return;
        }

        // Add result if all conditions match
        results.push({
            group: selectedGroup.displayName || "N/A",
            member: matchingMember.displayName || "N/A",
            memberMail: matchingMember.mail || "N/A",
            groupMail: selectedGroup.mail || "N/A",
            groupDescription: selectedGroup.description || "N/A",
            memberCount: members.value.length,
        });
    }

    // Handle other individual and combined filters below
    if (groupId && !userId && !mailEnabled) {
        const selectedGroup = allGroups.find(group => group.id === groupId);
        if (!selectedGroup) {
            alert("Invalid group selection.");
            return;
        }

        const members = await callGraphApi(`/groups/${groupId}/members?$select=id,displayName,mail,accountEnabled`);
        const disabledMembers = members.value.filter(member =>
            member["@odata.type"] === "#microsoft.graph.user" && member.accountEnabled === false
        );

        if (disabledMembers.length === 0) {
            alert("The selected group has no sign-in disabled users.");
            return;
        }

        disabledMembers.forEach(member => {
            results.push({
                group: selectedGroup.displayName,
                member: member.displayName || "N/A",
                memberMail: member.mail || "N/A",
                groupMail: selectedGroup.mail || "N/A",
                groupDescription: selectedGroup.description || "N/A",
                memberCount: members.value.length,
            });
        });
    }

    if (!groupId && userId && !mailEnabled) {
        const selectedUser = disabledUsers.find(user => user.id === userId);
        if (!selectedUser) {
            alert("Invalid member selection.");
            return;
        }

        const userGroups = [];
        for (const group of allGroups) {
            const members = await callGraphApi(`/groups/${group.id}/members?$select=id`);
            if (members.value.some(member => member.id === userId)) {
                userGroups.push(group);
            }
        }

        if (userGroups.length === 0) {
            alert("The selected user does not belong to any groups.");
            return;
        }

        userGroups.forEach(group => {
            results.push({
                group: group.displayName || "N/A",
                member: selectedUser.displayName || "N/A",
                memberMail: selectedUser.mail || "N/A",
                groupMail: group.mail || "N/A",
                groupDescription: group.description || "N/A",
                memberCount: group.memberCount || "N/A",
            });
        });
    }

    if (!groupId && !userId && mailEnabled) {
        const filteredGroups = allGroups.filter(group =>
            group.mailEnabled.toString() === mailEnabled
        );

        if (filteredGroups.length === 0) {
            alert("No groups match the mail-enabled filter.");
            return;
        }

        filteredGroups.forEach(group => {
            results.push({
                group: group.displayName || "N/A",
                member: "N/A",
                memberMail: "N/A",
                groupMail: group.mail || "N/A",
                groupDescription: group.description || "N/A",
                memberCount: group.memberCount || "N/A",
            });
        });
    }

    if (groupId && !userId && mailEnabled) {
        const selectedGroup = allGroups.find(group => group.id === groupId);

        if (!selectedGroup) {
            alert("Invalid group selection.");
            return;
        }

        if (selectedGroup.mailEnabled.toString() !== mailEnabled) {
            alert("The selected group does not match the mail-enabled filter.");
            return;
        }

        const members = await callGraphApi(`/groups/${groupId}/members?$select=id,displayName,mail,accountEnabled`);
        const disabledMembers = members.value.filter(member =>
            member["@odata.type"] === "#microsoft.graph.user" && member.accountEnabled === false
        );

        if (disabledMembers.length === 0) {
            alert("The selected group has no sign-in disabled users.");
            return;
        }

        disabledMembers.forEach(member => {
            results.push({
                group: selectedGroup.displayName,
                member: member.displayName || "N/A",
                memberMail: member.mail || "N/A",
                groupMail: selectedGroup.mail || "N/A",
                groupDescription: selectedGroup.description || "N/A",
                memberCount: members.value.length,
            });
        });
    }

    if (!groupId && userId && mailEnabled) {
        const selectedUser = disabledUsers.find(user => user.id === userId);

        if (!selectedUser) {
            alert("Invalid member selection.");
            return;
        }

        const userGroups = [];
        for (const group of allGroups) {
            if (group.mailEnabled.toString() !== mailEnabled) continue;

            const members = await callGraphApi(`/groups/${group.id}/members?$select=id`);
            if (members.value.some(member => member.id === userId)) {
                userGroups.push(group);
            }
        }

        if (userGroups.length === 0) {
            alert("The selected user does not belong to any groups that match the mail-enabled filter.");
            return;
        }

        userGroups.forEach(group => {
            results.push({
                group: group.displayName || "N/A",
                member: selectedUser.displayName || "N/A",
                memberMail: selectedUser.mail || "N/A",
                groupMail: group.mail || "N/A",
                groupDescription: group.description || "N/A",
                memberCount: group.memberCount || "N/A",
            });
        });
    }

    // Display results
    displayResults(results);
}



// Display Results
function displayResults(results) {
    const outputBody = document.getElementById("outputBody");
    if (results.length === 0) {
        alert("No matching results found.");
        outputBody.innerHTML = "";
        return;
    }
    outputBody.innerHTML = results.map(result => `
        <tr>
            <td>${result.group}</td>
            <td>${result.member}</td>
            <td>${result.memberMail}</td>
            <td>${result.groupMail}</td>
            <td>${result.groupDescription}</td>
            <td>${result.memberCount}</td>
        </tr>
    `).join("");
}

// Utility Functions

async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["User.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send"],
            account,
        });

        const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
            method,
            headers: {
                Authorization: `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            body: body ? JSON.stringify(body) : null, // Ensure body is correctly serialized
        });

        if (response.ok) {
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("application/json")) {
                return await response.json(); // Parse JSON response
            }
            return {}; // Return empty object for empty responses like 204 No Content
        } else {
            const errorText = await response.text();
            console.error(`Graph API Error (${response.status}):`, errorText);
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
    } catch (error) {
        console.error("Error in callGraphApi:", error);
        throw error;
    }
}



// Download Report as CSV
function downloadReportAsCSV() {
    const headers = ["Group", "Member Display Name", "Member Mail", "Group Mail", "Group Description", "Members"];
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data available to download.");
        return;
    }

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Sign-In_Disabled_Users_Report.csv";
    link.click();
}


// Mail Report to Admin
async function sendReportAsMail() {
    const adminEmail = document.getElementById("adminEmail").value;

    if (!adminEmail) {
        alert("Please provide an admin email.");
        return;
    }

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    if (!rows.length) {
        alert("No data to send via email.");
        return;
    }

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `
        <table border="1">
            <thead>
                <tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr>
            </thead>
            <tbody>${emailContent}</tbody>
        </table>
    `;

    const message = {
        message: {
            subject: "Sign-In Disabled Users Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: adminEmail } }],
        },
    };

    try {
        // POST request using the corrected callGraphApi function
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent successfully!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report. Please try again.");
    }
}
