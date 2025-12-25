/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(async () => {
    document
        .getElementById("generate")
        .addEventListener("click", generateBirthdays);
});

async function generateBirthdays() {
    const status = document.getElementById("status");
    status.textContent = "Signing in...";

    try {
        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true
        });

        status.textContent = "Generating birthdays...";

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

        if (!response.ok) throw new Error("Backend error");

        status.textContent = "Birthday events created";
    } catch (err) {
        console.error(err);
        status.textContent = "Something went wrong";
    }
}



