async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        // Prepare the request body
        const body = { year: new Date().getFullYear() };

        // Call your Azure Function anonymously
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

// Optional: attach the function to your button click
document.getElementById("generateButton").addEventListener("click", generateBirthdays);
