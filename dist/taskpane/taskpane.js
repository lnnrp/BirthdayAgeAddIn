async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        
        const body = { year: new Date().getFullYear() };

        const response = await fetch(
            "https://birthdaysync.azurewebsites.net/api/generate-birthdays",
            {
                method: "POST",
                headers: {
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
