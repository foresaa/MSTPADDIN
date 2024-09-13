// Define the function that is called when the button is clicked
function runPythonTask() {
  var xhr = new XMLHttpRequest();
  xhr.open("POST", "https://mstpaddin.onrender.com/load-current-project-data/", true);
  xhr.setRequestHeader("Content-Type", "application/json");

  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4 && xhr.status === 200) {
      console.log("Task executed successfully!");
    } else if (xhr.readyState === 4) {
      console.error("Error executing task.");
    }
  };

  // Sending request to FastAPI to trigger the task
  xhr.send(JSON.stringify({}));
}
