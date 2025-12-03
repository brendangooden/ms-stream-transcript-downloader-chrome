// Background service worker for MS Teams Transcript Downloader
// Intercepts network requests to capture transcript URLs

console.log('[MS Teams Transcript Downloader] Background script loaded');

let transcriptUrl = null;
let temporaryDownloadUrl = null;
let capturedRequests = new Map();

// Listen for web requests to capture the media API calls
chrome.webRequest.onCompleted.addListener(
  async (details) => {
    const url = details.url;
    
    // Look for the media API endpoint that returns transcript metadata
    if (url.includes('_api/v2.1/drives') && 
        url.includes('items/') && 
        url.includes('media') && 
        url.includes('transcripts')) {
      console.log('[Transcript Downloader] Captured media API URL:', url);
      
      // We can't directly access the response body in Manifest V3
      // We'll need to re-fetch it or have the content script send it to us
      capturedRequests.set(details.requestId, {
        url: url,
        timestamp: Date.now()
      });
    }
  },
  { urls: ["https://*.sharepoint.com/*", "https://*.office.com/*"] }
);

// Clean up old requests periodically
setInterval(() => {
  const now = Date.now();
  const maxAge = 300000; // 5 minutes
  
  for (const [id, request] of capturedRequests.entries()) {
    if (now - request.timestamp > maxAge) {
      capturedRequests.delete(id);
    }
  }
}, 60000); // Clean up every minute

// Handle messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {

  console.log('[Transcript Downloader] Received message:', request.action);

  if (request.action === 'setTranscriptMetadata') {
    // Content script sends us the transcript metadata it captured
    temporaryDownloadUrl = request.temporaryDownloadUrl;
    transcriptUrl = request.temporaryDownloadUrl;
    console.log('[Transcript Downloader] Stored transcript URL:', temporaryDownloadUrl);
    sendResponse({ success: true });
    return true;
  }

  if (request.action === 'getTranscriptData') {
    // Return the captured transcript URL
    sendResponse({
      transcriptUrl: temporaryDownloadUrl || transcriptUrl
    });
    return true;
  }

  if (request.action === 'getStatus') {
    sendResponse({
      hasTranscriptUrl: !!transcriptUrl,
      capturedRequestsCount: capturedRequests.size
    });
    return true;
  }
});

// Listen for extension installation or update
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('[Transcript Downloader] Extension installed');
  } else if (details.reason === 'update') {
    console.log('[Transcript Downloader] Extension updated');
  }
});
