document.addEventListener("DOMContentLoaded", function () {
    const form = document.querySelector("form");
    let userData = [];

    form.addEventListener("submit", function (event) {
        event.preventDefault();

        const name = document.getElementById("name").value.trim();
        const email = document.getElementById("email").value.trim();
        const message = document.getElementById("message").value.trim();

        if (name === "" || email === "" || message === "") {
            alert("Please fill in all fields.");
            return;
        }
        
        userData.push({ Name: name, Email: email, Message: message });

        const ws = XLSX.utils.json_to_sheet(userData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "User Data");

        XLSX.writeFile(wb, "UserData.xlsx");

        alert("Data saved to Excel file!");
        form.reset();
    });
});
