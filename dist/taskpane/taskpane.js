const statusElement = document.getElementById("status");

async function generateBirthdays() {
    try {
        statusElement.textContent = "Generating birthdays...";

        const body = { year: new Date().getFullYear() };

        const response = await fetch(
            "https://birthdaysync.azurewebsites.net/api/GenerateBirthdays",
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(body)
            }
        );

        if (response.ok) {
            statusElement.textContent = "Birthdays generated successfully!";
        } else {
            const text = await response.text();
            statusElement.textContent = `Error generating birthdays: ${text}`;
        }

    } catch (err) {
        statusElement.textContent = `Error: ${err.message}`;
    }
}

// Attach the function to the button
document.getElementById("generateButton").addEventListener("click", generateBirthdays);
