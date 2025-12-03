// Popup script for MS Teams Transcript Downloader extension

document.addEventListener('DOMContentLoaded', async () => {
  console.log('[Popup] Loading extension status');

  try {
    // Query the background script for status
    const response = await chrome.runtime.sendMessage({ action: 'getStatus' });

    if (response) {
      updateStatus(response);
    } else {
      showError('Could not connect to background script');
    }

    // Load saved default format
    chrome.storage.sync.get(['defaultFormat'], (result) => {
      const defaultFormat = result.defaultFormat || 'json';
      document.getElementById('defaultFormat').value = defaultFormat;
    });

    // Save format preference when changed
    document.getElementById('defaultFormat').addEventListener('change', (e) => {
      const format = e.target.value;
      chrome.storage.sync.set({ defaultFormat: format }, () => {
        console.log('[Popup] Default format saved:', format);
      });
    });

  } catch (error) {
    console.error('[Popup] Error loading status:', error);
    showError('Error loading extension status');
  }
});

function updateStatus(status) {
  const statusContainer = document.querySelector('.status');
  const statusIcon = document.querySelector('.status-icon');
  const statusText = document.getElementById('status-text');
  const transcriptStatus = document.getElementById('transcript-status');
  const requestsCount = document.getElementById('requests-count');

  // Update status indicator
  if (status.hasTranscriptUrl) {
    statusContainer.classList.remove('inactive');
    statusContainer.classList.add('active');
    statusIcon.classList.remove('inactive');
    statusIcon.classList.add('active');
    statusText.textContent = 'Extension is active and ready';
  } else {
    statusContainer.classList.remove('active');
    statusContainer.classList.add('inactive');
    statusIcon.classList.remove('active');
    statusIcon.classList.add('inactive');
    statusText.textContent = 'Waiting for transcript page...';
  }

  // Update info items
  transcriptStatus.textContent = status.hasTranscriptUrl ? 'Yes ✓' : 'No';
  requestsCount.textContent = status.capturedRequestsCount || 0;
}

function showError(message) {
  const statusContainer = document.querySelector('.status');
  const statusIcon = document.querySelector('.status-icon');
  const statusText = document.getElementById('status-text');

  statusContainer.classList.remove('active');
  statusContainer.classList.add('inactive');
  statusIcon.classList.remove('active');
  statusIcon.classList.add('inactive');
  statusText.textContent = message;
}
