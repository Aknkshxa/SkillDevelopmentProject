const fileForm = document.getElementById('file-form');
const fileInput = document.getElementById('file-input');
const subjectNameInput = document.getElementById('subject-name');

fileForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    formData.append('subjectName', subjectNameInput.value);

    fetch('/process', {
        method: 'POST',
        body: formData,
    })
    .then((response) => response.json())
    .then((data) => console.log(data))
    .catch((error) => console.error(error));
});