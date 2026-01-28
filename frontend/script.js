const form = document.getElementById("uploadForm");
const summary = document.getElementById("srsSummary");
const downloadLink = document.getElementById("downloadLink");

// Replace with your live backend URL
const BACKEND_URL = "https://your-backend-url.com";

form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const formData = new FormData(form);

    try {
        summary.textContent = "Generating test cases... Please wait.";
        downloadLink.style.display = "none";

        const res = await fetch(`${BACKEND_URL}/generate`, {
            method: "POST",
            body: formData
        });

        const data = await res.json();

        if (data.error) {
            summary.textContent = "Error: " + data.error;
            return;
        }

        // Show SRS summary
        summary.textContent = `Component: ${data.component}\n\nSRS Summary:\n${data.srs_summary}`;

        // Show download link
        downloadLink.href = `${BACKEND_URL}${data.download_url}`;
        downloadLink.style.display = "inline";
        downloadLink.textContent = "Download Excel";

    } catch (err) {
        console.error(err);
        summary.textContent = "Error connecting to backend.";
    }
});

