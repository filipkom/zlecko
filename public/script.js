const companies = ["Apple", "Microsoft", "Google", "Amazon", "Facebook"];
const suggestionsDiv = document.getElementById("suggestions");
const companyInput = document.getElementById("companyInput");

companyInput.addEventListener("input", function() {
    const input = this.value.toLowerCase();
    suggestionsDiv.innerHTML = "";
    
    if (input.length >= 3) {
        fetch('/api/firmy', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ nazwa: input }),
        })
        .then(response => response.json())
        .then(data => {
            const div = document.createElement("div");
            div.textContent = data.nazwa;
            div.onclick = function() {
                companyInput.value = data.nazwa;
                document.getElementById("companyAddress").value = data.adres;
                document.getElementById("companyNIP").value = data.nip;
                suggestionsDiv.innerHTML = "";
            };
            suggestionsDiv.appendChild(div);
        })
        .catch(error => console.error("Błąd:", error));
    }
});

function generateExcelFromTemplate() {
    const companyName = document.getElementById("companyInput").value;
    if (!companyName) {
        alert("Proszę wybrać nazwę firmy");
        return;
    }

    fetch('/api/generuj-excel', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ nazwa: companyName }),
    })
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = `Zlecenie_${companyName}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
    })
    .catch(error => {
        console.error("Błąd podczas generowania pliku Excel:", error);
        alert("Wystąpił błąd podczas generowania pliku Excel.");
    });
}
