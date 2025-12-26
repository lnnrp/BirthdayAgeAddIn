Office.onReady(() => {
    console.log("office ready");
    const btn = document.getElementById("generateTaskpaneButton");
    if (btn) {
        btn.onclick = generateBirthdays;
    } else {
        console.error("Button not found in taskpane HTML");
    }
});

async function generateBirthdays() {
    console.log("button clicked");
    const statusEl = document.getElementById("status");
    try {
        statusEl.innerText = "Generating birthdays...";

        const response = await fetch(
            "https://birthdaysync.azurewebsites.net/api/GenerateBirthdays",
            {
                method: "POST",
                headers: {
                    "x-api-key": "YOUR_STATIC_KEY",
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ year: new Date().getFullYear() })
            }
        );

        if (response.ok) {
            statusEl.innerText = await response.text();
        } else {
            const text = await response.text();
            statusEl.innerText = `Error: ${text}`;
        }
    } catch (err) {
        console.error(err);
        statusEl.innerText = `Error: ${err.message || err}`;
    }
}
