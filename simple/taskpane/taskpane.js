const BIRTHDAY_EXT_PROP =
    "String {00020329-0000-0000-C000-000000000046} Name BirthdayAgeAddIn";

Office.onReady(() => {
    const btn = document.getElementById("generateTaskpaneButton");
    btn.onclick = generateBirthdays;
});

async function generateBirthdays() {
    const status = document.getElementById("status");

    try {
        status.innerText = "Signing in…";

        const token = await OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true
        });

        const year = new Date().getFullYear();

        status.innerText = "Loading existing birthday events…";
        const existingEvents = await getExistingBirthdayEvents(token, year);

        status.innerText = "Loading contacts…";
        const contacts = await getAllContacts(token);

        status.innerText = "Creating / updating birthdays…";
        await createOrUpdateBirthdays(token, contacts, existingEvents, year);

        status.innerText = "✅ Birthdays generated successfully!";
    } catch (err) {
        console.error(err);
        status.innerText = `❌ ${err.message || err}`;
    }
}

// ---------------- GRAPH HELPERS ----------------

async function getAllContacts(token) {
    let url =
        "https://graph.microsoft.com/v1.0/me/contacts" +
        "?$select=displayName,birthday";

    const contacts = [];

    while (url) {
        const resp = await fetch(url, {
            headers: { Authorization: `Bearer ${token}` }
        });

        if (!resp.ok) throw new Error("Failed to load contacts");

        const data = await resp.json();
        contacts.push(...(data.value || []));
        url = data["@odata.nextLink"];
    }

    return contacts;
}

async function getExistingBirthdayEvents(token, year) {
    let url =
        "https://graph.microsoft.com/v1.0/me/events" +
        `?$expand=singleValueExtendedProperties(` +
        `$filter=id eq '${BIRTHDAY_EXT_PROP}')` +
        `&$filter=singleValueExtendedProperties/any(p:p/value eq '${year}')`;

    const events = {};

    while (url) {
        const resp = await fetch(url, {
            headers: { Authorization: `Bearer ${token}` }
        });

        if (!resp.ok) throw new Error("Failed to load existing events");

        const data = await resp.json();

        for (const evt of data.value || []) {
            events[evt.subject] = evt;
        }

        url = data["@odata.nextLink"];
    }

    return events;
}

async function createOrUpdateBirthdays(token, contacts, existingEvents, year) {
    for (const contact of contacts) {
        if (!contact.birthday) continue;

        const birth = new Date(contact.birthday);
        if (birth.getFullYear() <= 1) continue;

        let birthdayThisYear = new Date(
            year,
            birth.getMonth(),
            birth.getDate()
        );

        // Feb 29 handling
        if (
            birth.getMonth() === 1 &&
            birth.getDate() === 29 &&
            !isLeapYear(year)
        ) {
            birthdayThisYear = new Date(year, 1, 28);
        }

        const age = year - birth.getFullYear();
        const subject = `${contact.displayName} – ${age}`;

        const start = toUtc(birthdayThisYear);
        const end = toUtc(new Date(birthdayThisYear.getTime() + 3600000));

        if (existingEvents[subject]) {
            await updateEvent(token, existingEvents[subject].id, start, end);
        } else {
            await createEvent(token, subject, start, end, year);
        }
    }
}

async function createEvent(token, subject, start, end, year) {
    await fetch("https://graph.microsoft.com/v1.0/me/events", {
        method: "POST",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            subject,
            start,
            end,
            singleValueExtendedProperties: [
                {
                    id: BIRTHDAY_EXT_PROP,
                    value: year.toString()
                }
            ]
        })
    });
}

async function updateEvent(token, eventId, start, end) {
    await fetch(
        `https://graph.microsoft.com/v1.0/me/events/${eventId}`,
        {
            method: "PATCH",
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ start, end })
        }
    );
}

// ---------------- UTIL ----------------

function toUtc(date) {
    return {
        dateTime: date.toISOString().replace(".000Z", ""),
        timeZone: "UTC"
    };
}

function isLeapYear(year) {
    return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}
