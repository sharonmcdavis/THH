// Open the popup
function openDataPopup() {
    const popup = document.getElementById('data-popup');
    const dataContainer = document.getElementById('data-container');

    // Fetch the data for the current day
    fetch('/get-todays-data')
        .then(response => response.json())
        .then(data => {
            // Format the data into HTML
            let html = '';
            data.forEach(item => {
                html += `
                        <div class="time-period">
                            <h3>${item.time_period}</h3>
                            <p>${item.details}</p>
                        </div>
                    `;
            });

            // Insert the formatted data into the popup
            dataContainer.innerHTML = html;

            // Show the popup
            popup.classList.remove('hidden');
        })
        .catch(error => {
            console.error('Error fetching data:', error);
            dataContainer.innerHTML = '<p>Error loading data.</p>';
            popup.classList.remove('hidden');
        });
}

// Close the popup
function closeDataPopup() {
    const popup = document.getElementById('data-popup');
    popup.classList.add('hidden');
}