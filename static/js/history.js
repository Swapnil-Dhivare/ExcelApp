function addToTemplates(downloadId) {
    fetch(`/add_to_templates/${downloadId}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': document.querySelector('meta[name="csrf-token"]').content
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            // Show success message
            const toast = new bootstrap.Toast(document.getElementById('successToast'));
            toast.show();
            
            // Refresh page after delay
            setTimeout(() => {
                location.reload();
            }, 1500);
        } else {
            throw new Error(data.error || 'Unknown error occurred');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        // Show error message
        const errorToast = document.getElementById('errorToast');
        errorToast.querySelector('.toast-body').textContent = error.message;
        const toast = new bootstrap.Toast(errorToast);
        toast.show();
    });
}

// Add event listeners when document loads
document.addEventListener('DOMContentLoaded', function() {
    // Initialize tooltips
    const tooltips = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltips.forEach(tooltip => new bootstrap.Tooltip(tooltip));
});