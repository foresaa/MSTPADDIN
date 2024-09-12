function runPythonTask() {
    fetch("https://mstpaddin.onrender.com/load-current-project-data/", {
        method: "POST",
    })
    .then(response => response.json())
    .then(data => {
        console.log(data.message);  // Display the response message in the console
    })
    .catch(error => {
        console.error('Error:', error);
    });
}
