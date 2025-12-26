Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office is ready in Outlook.");

        // Attach click handler to button
        const btn = document.getElementById("generateButton");
        if (btn) {
            btn.addEventListener("click", generateBirthdays);
        }
    }
});

async function generateBirthdays() {
    try {
        console.log("Generate clicked");

        const body = { year: new Date().getFullYear() };

        const response = await fetch("https://birthdaysync.azurewebsites.net/api/GenerateBirthdays", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        });

        const statusEl = document.getElementById("status");

        if (response.ok) {
            console.log("Birthdays generated successfully!");
            if (statusEl) statusEl.textContent = "Birthdays generated successfully!";
        } else {
            const text = await response.text();
            console.error("Error generating birthdays:", text);
            if (statusEl) statusEl.textContent = `Error: ${text}`;
        }
    } catch (err) {
        console.error("Error in generateBirthdays:", err);
        const statusEl = document.getElementById("status");
        if (statusEl) statusEl.textContent = `Error: ${err}`;
    }
}
