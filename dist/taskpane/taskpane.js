Office.onReady(async () => {
    console.log("Office is ready");

    const generateButton = document.getElementById("generateButton");
    if (!generateButton) return console.error("Generate button not found!");

    generateButton.addEventListener("click", async () => {
        console.log("Generate button clicked");

        let token;
        try {
            // Attempt to get an access token
            token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
            console.log("Access token obtained:", token);
        } catch (err) {
            console.error("Failed to get access token:", err);
            alert("Failed to get access token. Check console for details.");
            return;
        }

        try {
            // Example API call - replace with your endpoint
            const response = await fetch("https://birthdaysync.azurewebsites.net/api/generate-birthdays", {
                method: "GET", // or "POST"
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            });

            console.log("Fetch response status:", response.status);

            if (!response.ok) {
                const text = await response.text();
                console.error("Fetch failed with status", response.status, "and response:", text);
                alert(`Fetch failed: ${response.status}`);
                return;
            }

            const data = await response.json();
            console.log("Fetch succeeded, data:", data);
            alert("Fetch succeeded! Check console for data.");
        } catch (err) {
            console.error("Fetch threw an error:", err);
            alert("Fetch error occurred. Check console for details.");
        }
    });
});
