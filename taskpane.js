function checkSpelling() {
  Office.context.mailbox.item.body.getAsync("text", function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const text = result.value;
      const apiKey = "YOUR_LANGUAGETOOL_API_KEY";
      const url = "https://api.languagetool.org/v2/check";
      const params = {
        text: text,
        language: "auto"
      };

      fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${apiKey}`
        },
        body: JSON.stringify(params)
      })
      .then(response => response.json())
      .then(data => {
        const resultsDiv = document.getElementById("results");
        resultsDiv.innerHTML = "";
        data.matches.forEach(match => {
          const p = document.createElement("p");
          p.textContent = `Error: ${match.message} (suggestions: ${match.replacements.map(r => r.value).join(", ")})`;
          resultsDiv.appendChild(p);
        });
      })
      .catch(error => console.error("Error:", error));
    }
  });
}
