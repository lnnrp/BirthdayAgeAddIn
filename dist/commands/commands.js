/* global Office */

Office.onReady(() => {});

// This function will be triggered by the ribbon button
async function generateBirthdays(event) {
    try {
        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

        const year = new Date().getFullYear();

        const response = await fetch(
            "https://birthdaysync.azurewebsites.net/api/generate-birthdays",
            {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ year })
            }
        );

        if (!response.ok) throw new Error("Azure Function call failed");

        console.log("Birthday events generated successfully");
    } catch (err) {
        console.error(err);
    } finally {
        // Notify Outlook the action finished
        event.completed();
    }
}

// Make the function globally available
window.generateBirthdays = generateBirthdays;
