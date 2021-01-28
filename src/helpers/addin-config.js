function getConfig() {
    var config = {};
  
    config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
    config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');
  
    return config;
  }
  
  function setConfig(config, callback) {
    Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
    Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);
  
    Office.context.roamingSettings.saveAsync(callback);
  }

  function getGist(gistId, callback) {
    var requestUrl = 'https://api.github.com/gists/' + gistId;
  
    $.ajax({
      url: requestUrl,
      dataType: 'json'
    }).done(function(gist){
      callback(gist);
    }).fail(function(error){
      callback(null, error);
    });
  }
  
  function buildBodyContent(gist, callback) {
    // Find the first non-truncated file in the gist
    // and use it.
    for (var filename in gist.files) {
      if (gist.files.hasOwnProperty(filename)) {
        var file = gist.files[filename];
        if (!file.truncated) {
          // We have a winner.
          switch (file.language) {
            case 'HTML':
              // Insert as-is.
              callback(file.content);
              break;
            case 'Markdown':
              // Convert Markdown to HTML.
              var converter = new showdown.Converter();
              var html = converter.makeHtml(file.content);
              callback(html);
              break;
            default:
              // Insert contents as a <code> block.
              var codeBlock = '<pre><code>';
              codeBlock = codeBlock + file.content;
              codeBlock = codeBlock + '</code></pre>';
              callback(codeBlock);
          }
          return;
        }
      }
    }
    callback(null, 'No suitable file found in the gist');
  }