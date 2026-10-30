// Voice Input Support
const voiceBtns = document.querySelectorAll('.voice-btn');
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition = null;
let isSpeechSupported = false;

if (SpeechRecognition) {
    try {
        recognition = new SpeechRecognition();
        recognition.continuous = false;
        recognition.interimResults = false;
        isSpeechSupported = true;
    } catch (e) {
        console.error('Speech recognition initialization failed:', e);
    }
}

if (isSpeechSupported && recognition) {
    voiceBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-target');
            const targetInput = document.getElementById(targetId);
            const originalIcon = btn.innerHTML;
            
            btn.innerHTML = '<span class="material-icons align-middle spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>';
            btn.disabled = true;
            
            recognition.lang = document.getElementById('language').value === 'Hindi' ? 'hi-IN' : 'en-US';
            recognition.start();

            recognition.onresult = (event) => {
                const transcript = event.results[0][0].transcript;
                targetInput.value = transcript;
                btn.innerHTML = originalIcon;
                btn.disabled = false;
            };

            recognition.onerror = () => {
                btn.innerHTML = originalIcon;
                btn.disabled = false;
                alert('Voice recognition error. Please try again.');
            };

            recognition.onspeechend = () => {
                recognition.stop();
                btn.innerHTML = originalIcon;
                btn.disabled = false;
            };
        });
    });
} else {
    voiceBtns.forEach(btn => btn.style.display = 'none');
}

// Form Submission & Download Tracking
document.getElementById('pptForm').addEventListener('submit', function (e) {
    // Clear previous download cookies
    document.cookie = "ppt_downloaded=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
    document.cookie = "ppt_error=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";

    const statusText = document.getElementById('statusText');
    const timer = document.getElementById('timer');
    const downloadButtons = document.getElementById('downloadButtons');

    statusText.textContent = 'Generating masterpiece...';
    statusText.className = 'text-muted';
    timer.classList.remove('d-none');
    downloadButtons.classList.add('d-none');

    let seconds = 0;
    const timerInterval = setInterval(() => {
        seconds++;
        timer.textContent = `${seconds}s`;
    }, 1000);

    const checkInterval = setInterval(() => {
        // Parse current cookies
        const cookies = document.cookie.split(';').reduce((acc, c) => {
            const [key, val] = c.trim().split('=');
            if (key) acc[key] = decodeURIComponent(val || '');
            return acc;
        }, {});

        if (cookies.ppt_downloaded === 'true') {
            cleanup();
            statusText.textContent = 'Success! Your masterpiece is downloaded.';
            statusText.className = 'text-success';
        } else if (cookies.ppt_error) {
            const errorMsg = cookies.ppt_error;
            cleanup();
            statusText.textContent = `Error: ${errorMsg}`;
            statusText.className = 'text-danger';
        }
    }, 500);

    // Keep the UI wait aligned with the server's 18-second Gemini request limit.
    const fallbackTimeout = setTimeout(() => {
        cleanup();
        statusText.textContent = 'Generation timed out. Please try again with fewer slides.';
        statusText.className = 'text-danger';
    }, 20000);

    function cleanup() {
        clearInterval(timerInterval);
        clearInterval(checkInterval);
        clearTimeout(fallbackTimeout);
        document.cookie = "ppt_downloaded=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
        document.cookie = "ppt_error=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
        timer.classList.add('d-none');
    }
});
