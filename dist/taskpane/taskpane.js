Office.onReady(() => {
    console.log("Office ready in taskpane");
});

async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

        console.log("Token acquired");

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

        if (!response.ok) {
            throw new Error(await response.text());
        }

        alert("Birthday events generated successfully 🎉");
    } catch (err) {
        console.error(err);
        alert("Error: " + err.message);
    }
}

window.generateBirthdays = generateBirthdays;
