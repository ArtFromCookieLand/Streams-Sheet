/**
 * Service for interacting with the Apify API.
 */
var ApifyService = {
  
  /**
   * Starts an Actor run with specific input configuration.
   * @param {string} actorId - The ID of the actor (e.g., 'username/spotify-scraper')
   * @param {string} token - Your Apify API Token
   * @param {Object} actorInput - The JSON input for the actor (e.g., URLs to scrape)
   */
  startActor: function(actorId, token, actorInput) {
    const url = `https://api.apify.com/v2/acts/${actorId}/runs?token=${token}`;
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(actorInput) // Send the input configuration
    };
    
    console.log(`Starting Apify Actor: ${actorId}`);
    const response = UrlFetchApp.fetch(url, options);
    const run = JSON.parse(response.getContentText()).data;
    
    return run;
  },

  /**
   * Polls the run status until it is finished.
   */
  waitForCompletion: function(runId, actorId, token) {
    const url = `https://api.apify.com/v2/acts/${actorId}/runs/${runId}?token=${token}`;
    
    console.log(`Polling status for run: ${runId}...`);
    
    // Poll every 5 seconds for up to 5 minutes (60 attempts)
    for (let i = 0; i < 60; i++) {
      const response = UrlFetchApp.fetch(url);
      const status = JSON.parse(response.getContentText()).data.status;
      
      if (status === 'SUCCEEDED') {
        console.log('Actor run SUCCEEDED.');
        return true;
      } 
      if (status === 'FAILED' || status === 'ABORTED') {
        throw new Error('Apify Actor run failed or was aborted.');
      }
      
      Utilities.sleep(5000); 
    }
    
    throw new Error('Timed out waiting for Apify Actor to finish.');
  },

  /**
   * Fetches the dataset items.
   */
  fetchDataset: function(datasetId, token) {
    const url = `https://api.apify.com/v2/datasets/${datasetId}/items?token=${token}&format=json`;
    console.log(`Fetching dataset: ${datasetId}`);
    const response = UrlFetchApp.fetch(url);
    return JSON.parse(response.getContentText());
  }
};