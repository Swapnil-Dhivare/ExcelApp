// Function to get CSRF token
function getCSRFToken() {
    return document.querySelector('meta[name="csrf-token"]').getAttribute('content');
}

// Add CSRF token to all AJAX requests
document.addEventListener('DOMContentLoaded', function() {
    // For jQuery AJAX
    if (typeof $.ajaxSetup === 'function') {
        $.ajaxSetup({
            beforeSend: function(xhr, settings) {
                if (!/^(GET|HEAD|OPTIONS|TRACE)$/i.test(settings.type) && !this.crossDomain) {
                    xhr.setRequestHeader("X-CSRFToken", getCSRFToken());
                }
            }
        });
    }
    
    // For fetch API
    const originalFetch = window.fetch;
    window.fetch = function(url, options = {}) {
        if (options.method && options.method.toUpperCase() !== 'GET') {
            if (!options.headers) options.headers = {};
            options.headers['X-CSRFToken'] = getCSRFToken();
        }
        return originalFetch(url, options);
    };
    
    // For form submissions
    document.querySelectorAll('form[action="/add_sheet"]').forEach(form => {
        if (!form.querySelector('input[name="csrf_token"]')) {
            const csrfInput = document.createElement('input');
            csrfInput.type = 'hidden';
            csrfInput.name = 'csrf_token';
            csrfInput.value = getCSRFToken();
            form.appendChild(csrfInput);
        }
    });
});