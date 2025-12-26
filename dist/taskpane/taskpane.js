Office.onReady(() => {
    // Ensure Office is ready before adding event listeners
    document.getElementById("generateButton").onclick = generateBirthdays;
});

async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        let token;

        // Use OfficeRuntime.auth for modern SSO (OWA + Desktop)
        if (Office.context.requirements.isSetSupported("IdentityAPI", 1.3)) {
            token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
            console.log("Got token via OfficeRuntime.auth");
        } else {
            // Fallback for older Outlook desktop clients
            token = await new Promise((resolve, reject) => {
                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
                    if (result.status === "succeeded") {
                        resolve(result.value);
                    } else {
                        reject(result.error);
                    }
                });
            });
            console.log("Got token via getCallbackTokenAsync");
        }

        // Prepare body
        const body = { year: new Date().getFullYear() };

        // Call your Azure Function
        const response = await fetch("https://birthdaysync.azurewebsites.net/api/GenerateBirthdays", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        });

        if (response.ok) {
            console.log("Birthdays generated successfully!");
            document.getElementById("status").innerText = "Birthdays generated successfully!";
        } else {
            const text = await response.text();
            console.error("Error generating birthdays:", text);
            document.getElementById("status").innerText = `Error generating birthdays: ${text}`;
        }
    } catch (err) {
        console.error("Error in generateBirthdays:", err);
        document.getElementById("status").innerText = `Error: ${err.message || err}`;
    }
}
