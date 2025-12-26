async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        let token;

        // Use OfficeRuntime API (works in desktop, OWA, mobile)
        if (OfficeRuntime && OfficeRuntime.auth && Office.context.requirements.isSetSupported('IdentityAPI', 1.3)) {
            token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        } else {
            console.error("IdentityAPI not supported on this platform.");
            return;
        }

        const body = { year: new Date().getFullYear() };

        const response = await fetch(
            "https://birthdaysync.azurewebsites.net/api/generate-birthdays",
            {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(body)
            }
        );

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
