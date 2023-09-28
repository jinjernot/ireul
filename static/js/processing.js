function showProcessingMessage() {
  var modal = document.getElementById("processingModal");
  modal.style.display = "block";

  // Send an asynchronous request to the server
  fetch('/upload-file-data')
    .then(response => {
      if (response.ok) {
        return response.text();
      } else {
        throw new Error('Network response was not ok.');
      }
    })
    .then(data => {
      // Handle the data returned from the server
      console.log(data);
      // Hide the processing message modal
      modal.style.display = "none";
    })
    .catch(error => {
      // Handle any errors that occurred during the request
      console.error('Error:', error);
      // Hide the processing message modal
      modal.style.display = "none";
    });

  return true;
}
