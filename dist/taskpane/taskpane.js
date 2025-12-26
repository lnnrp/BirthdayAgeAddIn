async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        let token;

        // Use the new OfficeRuntime API for web/OWA
        if (Office.context.requirements.isSetSupported('IdentityAPI', 1.3)) {
            token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        } else {
            // Fallback for older desktop versions
            token = await new Promise((resolve, reject) => {
                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function(result) {
                    if (result.status === "succeeded") {
                        resolve(result.value);
                    } else {
                        reject(result.error);
                    }
                });
            });
        }

        // Prepare the request body
        const body = { year: new Date().getFullYear() };

        // Call your Azure Function
        const response = await fetch("https://birthdaysync.azurewebsites.net/api/generate-birthdays", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        });

        if (response.ok) {
            console.log("Birthdays generated successfully!");
        } else {
            const text = await response.text();
            console.error("Error generating birthdays:", text);
        }

    } catch (err) {
        console.error("Error in generateBirthdays:", err);
    }
}
