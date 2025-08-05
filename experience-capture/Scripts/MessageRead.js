
(function () {
    "use strict";

    // Configuration loaded from build-time
    const APP_CONFIG = {
  "appInfo": {
    "id": "0442beaa-afd6-4434-bc56-8dadf65db0aa",
    "name": "Experience Capture",
    "description": "Updates Experience directly from email inquiries",
    "buttonLabel": "Experience Capture",
    "version": "1.0.0.0",
    "provider": "HSO"
  },
  "deployment": {
    "webServerUrl": "https://lemon-rock-0d0849e10.1.azurestaticapps.net",
    "environment": "production",
    "comments": "Update webServerUrl to match your actual web server where add-in files will be hosted"
  },
  "canvasApp": {
    "url": "https://apps.powerapps.com/play/e/4b3631c7-c916-4ebc-935c-bfa92317ad03/a/2876bcfe-ebda-4542-ac05-62ee47efa210",
    "height": "700px"
  },
  "emailContext": {
    "fields": [
      "subject",
      "sender",
      "recipients",
      "body"
    ],
    "parameterFormat": "urlParams"
  },
  "activationRules": {
    "itemTypes": [
      "Message"
    ],
    "formTypes": [
      "Read"
    ],
    "customRules": []
  },
  "bodyProcessing": {
    "enabled": true,
    "extractionRules": [
      {
        "name": "recordId",
        "description": "Extract Dataverse record ID from URL",
        "method": "regex",
        "pattern": "[&](?:amp;)id=([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})",
        "group": 1,
        "required": true
      },
      {
        "name": "entityType",
        "description": "Extract entity type from URL",
        "method": "regex",
        "pattern": "[&](?:amp;)etn=([a-zA-Z_]+)",
        "group": 1,
        "required": false,
        "defaultValue": "unknown"
      },
      {
        "name": "appId",
        "description": "Extract app ID from URL",
        "method": "regex",
        "pattern": "[&](?:amp;)appid=([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})",
        "group": 1,
        "required": false
      },
      {
        "name": "dataverseUrl",
        "description": "Extract full Dataverse URL",
        "method": "regex",
        "pattern": "(https://[a-zA-Z0-9.-]+\\.crm[0-9]*\\.dynamics\\.com/[^\\s]+)",
        "group": 1,
        "required": false
      },
      {
        "name": "orgName",
        "description": "Extract organization name from URL",
        "method": "regex",
        "pattern": "https://([a-zA-Z0-9.-]+)\\.crm[0-9]*\\.dynamics\\.com",
        "group": 1,
        "required": false
      },
      {
        "name": "hasDataverseLink",
        "description": "Check if email contains Dataverse link",
        "method": "keywords",
        "keywords": {
          "true": [
            "crm.dynamics.com",
            "dynamics.com/main.aspx"
          ],
          "false": []
        },
        "defaultValue": "false"
      },
      {
        "name": "actionType",
        "description": "Determine what action is mentioned with the record",
        "method": "keywords",
        "keywords": {
          "review": [
            "review",
            "check",
            "look at",
            "examine"
          ],
          "update": [
            "update",
            "modify",
            "change",
            "edit"
          ],
          "approve": [
            "approve",
            "approval",
            "sign off"
          ],
          "follow_up": [
            "follow up",
            "follow-up",
            "contact",
            "call"
          ],
          "urgent": [
            "urgent",
            "asap",
            "immediately",
            "critical"
          ]
        },
        "defaultValue": "review"
      }
    ],
    "textCleaning": {
      "removeHtmlTags": true,
      "removeExtraWhitespace": true,
      "removeEmailHeaders": false,
      "removeSignatures": false
    }
  },
  "ui": {
    "showDebugInfo": false,
    "loadingMessage": "Loading Experience Capture...",
    "errorRetryEnabled": true
  }
};

    // Initialize Office Add-in
    Office.onReady(function (reason) {
        document.addEventListener('DOMContentLoaded', function() {
            console.log('Experience Capture - Office Add-in initialized');
            initializeCanvasApp();
        });
        
        // If DOM is already loaded
        if (document.readyState === 'loading') {
            // DOM not ready yet
        } else {
            console.log('Experience Capture - Office Add-in initialized');
            initializeCanvasApp();
        }
    });

    // Main initialization function
    function initializeCanvasApp() {
        try {
            showLoading(APP_CONFIG.ui.loadingMessage);
            extractEmailContext();
        } catch (error) {
            console.error('Initialization error:', error);
            showError('Failed to initialize add-in: ' + error.message);
        }
    }

    // Extract email context based on configuration - UPDATED FOR HTML PROCESSING
    function extractEmailContext() {
        console.log("Extracting email context...");

        try {
            const item = Office.context.mailbox.item;

            if (!item) {
                throw new Error("No email item found");
            }

            console.log("Email item found, collecting context...");

            // Build basic email context
            const emailContext = {
                subject: item.subject || '',
                sender: item.from ? item.from.emailAddress : '',
                senderName: item.from ? item.from.displayName : '',
                recipients: item.to ? item.to.map(r => r.emailAddress).join(';') : '',
                attachmentCount: item.attachments ? item.attachments.length : 0,
                hasAttachments: item.attachments && item.attachments.length > 0
            };

            // Check if body processing is needed
            if (APP_CONFIG.emailContext.fields.includes('body') &&
                APP_CONFIG.bodyProcessing &&
                APP_CONFIG.bodyProcessing.enabled) {

                // Get BOTH text and HTML body for comprehensive processing
                let bodiesReceived = 0;
                let textBody = '';
                let htmlBody = '';

                function processWhenBothBodiesReceived() {
                    bodiesReceived++;
                    if (bodiesReceived === 2) {
                        console.log("Both email bodies extracted");
                        console.log("Text body length:", textBody.length);
                        console.log("HTML body length:", htmlBody.length);

                        // Process both versions with HTML-aware logic
                        const processedBody = processEmailBodyAdvanced(textBody, htmlBody, APP_CONFIG.bodyProcessing);
                        emailContext.processedBody = processedBody;

                        // Build Canvas App URL with processed data
                        const canvasUrl = buildCanvasAppUrl(emailContext);
                        showCanvasApp(canvasUrl, emailContext);
                    }
                }

                // Get text body
                item.body.getAsync(Office.CoercionType.Text, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        textBody = result.value;
                        console.log("Text body extracted successfully");
                    } else {
                        console.error('Failed to get text body:', result.error);
                        textBody = '';
                    }
                    processWhenBothBodiesReceived();
                });

                // Get HTML body
                item.body.getAsync(Office.CoercionType.Html, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        htmlBody = result.value;
                        console.log("HTML body extracted successfully");
                    } else {
                        console.error('Failed to get HTML body:', result.error);
                        htmlBody = '';
                    }
                    processWhenBothBodiesReceived();
                });

            } else {
                // No body processing needed, load immediately
                const canvasUrl = buildCanvasAppUrl(emailContext);
                showCanvasApp(canvasUrl, emailContext);
            }

        } catch (error) {
            console.error("Error in extractEmailContext:", error);
            showError("Failed to extract email context: " + error.message);
        }
    }

    // NEW: Advanced processing that handles both text and HTML
    function processEmailBodyAdvanced(textBody, htmlBody, processingConfig) {
        try {
            console.log('Processing email body (advanced) with both text and HTML');

            // First, extract URLs from HTML before any cleaning
            const urlsFromHtml = extractUrlsFromHtml(htmlBody);
            console.log('URLs extracted from HTML:', urlsFromHtml);

            // Clean the text body
            let cleanedTextBody = textBody;
            if (processingConfig.textCleaning) {
                cleanedTextBody = cleanEmailText(textBody, processingConfig.textCleaning);
            }

            // Combine text with extracted URLs for processing
            const combinedText = cleanedTextBody + '\n\n' + urlsFromHtml.join('\n');

            const extractedData = {
                originalLength: textBody.length,
                htmlLength: htmlBody.length,
                cleanedLength: cleanedTextBody.length,
                extractedUrlsCount: urlsFromHtml.length,
                cleanedText: cleanedTextBody.substring(0, 200),
                extractedUrls: urlsFromHtml,
                extractedFields: {},
                processingLog: []
            };

            // Apply extraction rules to the combined text
            if (processingConfig.extractionRules) {
                processingConfig.extractionRules.forEach(rule => {
                    try {
                        console.log('Applying extraction rule:', rule.name);
                        const extractedValue = applyExtractionRule(combinedText, rule);
                        extractedData.extractedFields[rule.name] = extractedValue;
                        extractedData.processingLog.push({
                            rule: rule.name,
                            value: extractedValue,
                            success: extractedValue !== null
                        });
                        console.log(`Extracted ${rule.name}:`, extractedValue);
                    } catch (error) {
                        console.error(`Error applying rule ${rule.name}:`, error);
                        extractedData.extractedFields[rule.name] = rule.defaultValue || null;
                        extractedData.processingLog.push({
                            rule: rule.name,
                            error: error.message,
                            success: false
                        });
                    }
                });
            }

            console.log('Advanced body processing complete:', extractedData);
            return extractedData;

        } catch (error) {
            console.error('Error processing email body (advanced):', error);
            return {
                error: error.message,
                originalLength: textBody.length,
                htmlLength: htmlBody.length,
                extractedFields: {},
                processingLog: [{ error: error.message, success: false }]
            };
        }
    }

    // NEW: Extract URLs from HTML content
    function extractUrlsFromHtml(htmlContent) {
        const urls = [];

        if (!htmlContent) return urls;

        try {
            // Extract href attributes from <a> tags
            const hrefRegex = /<a[^>]+href\s*=\s*["']([^"']+)["'][^>]*>/gi;
            let match;

            while ((match = hrefRegex.exec(htmlContent)) !== null) {
                const url = match[1];
                // Only include URLs that look like Dataverse URLs
                if (url.includes('crm') && url.includes('dynamics.com')) {
                    urls.push(url);
                }
            }

            // Also look for URLs in the text content (not in href attributes)
            const urlRegex = /https?:\/\/[^\s<>"']+crm[^\s<>"']*/gi;
            let urlMatch;

            while ((urlMatch = urlRegex.exec(htmlContent)) !== null) {
                const url = urlMatch[0];
                if (!urls.includes(url)) {
                    urls.push(url);
                }
            }

            console.log('Extracted URLs from HTML:', urls);
            return urls;

        } catch (error) {
            console.error('Error extracting URLs from HTML:', error);
            return [];
        }
    }

    // Clean email text according to configuration - UPDATED
    function cleanEmailText(text, cleaningConfig) {
        let cleaned = text;

        // UPDATED: Extract URLs BEFORE removing HTML tags
        const extractedUrls = [];
        if (cleaningConfig.removeHtmlTags) {
            // Extract href attributes from <a> tags before removing HTML
            const hrefRegex = /<a[^>]+href\s*=\s*["']([^"']+)["'][^>]*>/gi;
            let match;
            while ((match = hrefRegex.exec(text)) !== null) {
                extractedUrls.push(match[1]);
            }

            // Replace the <a> tag with: linkText followed by the actual URL
            cleaned = cleaned.replace(/<a[^>]+href\s*=\s*["']([^"']+)["'][^>]*>([^<]*)<\/a>/gi,
                function (fullMatch, url, linkText) {
                    return linkText + ' ' + url + ' ';
                }
            );

            // Now remove remaining HTML tags
            cleaned = cleaned.replace(/<[^>]*>/g, ' ');

            // If we found URLs but they might have been lost, add them back
            if (extractedUrls.length > 0) {
                cleaned += '\n\nExtracted URLs:\n' + extractedUrls.join('\n');
            }
        }

        if (cleaningConfig.removeEmailHeaders) {
            // Remove common email headers
            cleaned = cleaned.replace(/^(From:|To:|Sent:|Subject:).*$/gm, '');
        }

        if (cleaningConfig.removeSignatures) {
            // Remove common signature patterns
            cleaned = cleaned.replace(/^--\s*$/gm, '<!-- SIGNATURE_BREAK -->');
            const parts = cleaned.split('<!-- SIGNATURE_BREAK -->');
            cleaned = parts[0]; // Keep only the first part (before signature)
        }

        if (cleaningConfig.removeExtraWhitespace) {
            // Remove extra whitespace
            cleaned = cleaned.replace(/\s+/g, ' ').trim();
        }

        return cleaned;
    }

    // Apply individual extraction rule
    function applyExtractionRule(text, rule) {
        switch (rule.method) {
            case 'regex':
                return extractByRegex(text, rule);
            case 'keywords':
                return extractByKeywords(text, rule);
            case 'sentences':
                return extractSentences(text, rule);
            case 'context':
                return extractContext(text, rule);
            default:
                console.warn('Unknown extraction method:', rule.method);
                return rule.defaultValue || null;
        }
    }

    // Extract using regex pattern
    function extractByRegex(text, rule) {
        try {
            const regex = new RegExp(rule.pattern, 'gim');
            const matches = text.match(regex);
            
            if (matches && matches.length > 0) {
                // If we want a specific group, extract it from the first match
                if (rule.group !== undefined) {
                    const specificMatch = text.match(new RegExp(rule.pattern, 'im'));
                    if (specificMatch && specificMatch[rule.group]) {
                        return specificMatch[rule.group].trim();
                    }
                } else {
                    // Return the first full match
                    return matches[0].trim();
                }
            }
            
            return rule.defaultValue || null;
        } catch (error) {
            console.error('Regex extraction error for rule', rule.name, ':', error);
            return rule.defaultValue || null;
        }
    }

    // Extract using keyword matching
    function extractByKeywords(text, rule) {
        const lowerText = text.toLowerCase();
        
        for (const [category, keywords] of Object.entries(rule.keywords)) {
            for (const keyword of keywords) {
                if (lowerText.includes(keyword.toLowerCase())) {
                    return category;
                }
            }
        }
        
        return rule.defaultValue || null;
    }

    // Extract key sentences
    function extractSentences(text, rule) {
        try {
            // Split into sentences
            const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 20);
            
            let selectedSentences = sentences;
            
            // Skip common phrases if configured
            if (rule.skipCommonPhrases) {
                const commonPhrases = [
                    'thank you', 'best regards', 'sincerely', 'kind regards',
                    'please let me know', 'if you have any questions'
                ];
                
                selectedSentences = sentences.filter(sentence => {
                    const lower = sentence.toLowerCase();
                    return !commonPhrases.some(phrase => lower.includes(phrase));
                });
            }
            
            // Take first N sentences
            const maxSentences = rule.maxSentences || 3;
            return selectedSentences
                .slice(0, maxSentences)
                .map(s => s.trim())
                .join('. ') + '.';
                
        } catch (error) {
            console.error('Sentence extraction error:', error);
            return rule.defaultValue || '';
        }
    }

    // Extract context around a pattern
    function extractContext(text, rule) {
        try {
            const regex = new RegExp(rule.pattern, 'i');
            const match = text.search(regex);
            
            if (match !== -1) {
                const beforeChars = rule.beforeChars || 50;
                const afterChars = rule.afterChars || 50;
                
                const start = Math.max(0, match - beforeChars);
                const urlMatch = text.match(regex);
                const urlLength = urlMatch ? urlMatch[0].length : 0;
                const end = Math.min(text.length, match + urlLength + afterChars);
                
                return text.substring(start, end).trim();
            }
            
            return rule.defaultValue || null;
        } catch (error) {
            console.error('Context extraction error:', error);
            return rule.defaultValue || null;
        }
    }

    // Build Canvas App URL with extracted data - UPDATED
    function buildCanvasAppUrl(emailContext) {
        try {
            const url = new URL(APP_CONFIG.canvasApp.url);
            
            // Add basic email metadata
            url.searchParams.append('subject', emailContext.subject || '');
            url.searchParams.append('sender', emailContext.sender || '');
            url.searchParams.append('senderName', emailContext.senderName || '');
            url.searchParams.append('timestamp', new Date().getTime().toString());
            
            // Add processed body data if available
            if (emailContext.processedBody) {
                const processed = emailContext.processedBody;
                
                // Add metadata about processing
                url.searchParams.append('originalLength', processed.originalLength.toString());
                url.searchParams.append('cleanedLength', processed.cleanedLength.toString());
                url.searchParams.append('cleanedText', processed.cleanedText);
                
                // NEW: Add HTML processing metadata
                if (processed.htmlLength !== undefined) {
                    url.searchParams.append('htmlLength', processed.htmlLength.toString());
                    url.searchParams.append('extractedUrlsCount', processed.extractedUrlsCount.toString());
                }
                
                // Add all extracted fields as URL parameters
                Object.keys(processed.extractedFields).forEach(fieldName => {
                    const value = processed.extractedFields[fieldName];
                    if (value !== null && value !== undefined) {
                        url.searchParams.append(fieldName, value.toString());
                    }
                });
                
                console.log('Extracted fields added to URL:', processed.extractedFields);
            }
            
            const finalUrl = url.toString();
            console.log('Built Canvas App URL:', finalUrl);
            
            return finalUrl;
            
        } catch (error) {
            console.error('Error building Canvas App URL:', error);
            return APP_CONFIG.canvasApp.url;
        }
    }

    // Show Canvas App with extracted data
    function showCanvasApp(url, emailContext) {
        console.log("Loading Canvas App:", url);
        hideAllContainers();
        
        const canvasContainer = document.getElementById('canvas-container');
        if (canvasContainer) {
            canvasContainer.style.display = 'block';
        }
        
        const iframe = document.getElementById('canvas-iframe');
        if (iframe) {
            iframe.src = url;
            
            // Show loading overlay
            const loadingOverlay = document.getElementById('iframe-loading');
            if (loadingOverlay) {
                loadingOverlay.style.display = 'flex';
            }
            
            // Handle iframe load
            iframe.onload = function() {
                console.log("Canvas App iframe loaded with extracted data");
                if (loadingOverlay) {
                    loadingOverlay.style.display = 'none';
                }
            };
            
            iframe.onerror = function() {
                console.error("Canvas App iframe failed to load");
                if (loadingOverlay) {
                    loadingOverlay.style.display = 'none';
                }
                showError('Failed to load Canvas App');
            };
        }
    }

    // UI Helper Functions - VANILLA JS
    function showLoading(message) {
        hideAllContainers();
        const loadingContainer = document.getElementById('loading-container');
        if (loadingContainer) {
            loadingContainer.style.display = 'block';
        }
        const loadingMessage = document.getElementById('loading-message');
        if (loadingMessage) {
            loadingMessage.textContent = message;
        }
    }

    function showError(message) {
        hideAllContainers();
        const errorContainer = document.getElementById('error-container');
        if (errorContainer) {
            errorContainer.style.display = 'block';
        }
        const errorMessage = document.getElementById('error-message');
        if (errorMessage) {
            errorMessage.textContent = message;
        }
    }

    function hideAllContainers() {
        const containers = ['loading-container', 'error-container', 'canvas-container'];
        containers.forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = 'none';
            }
        });
    }

    // Event Handlers - VANILLA JS
    document.addEventListener('DOMContentLoaded', function() {
        // Retry button
        const retryBtn = document.getElementById('retry-btn');
        if (retryBtn) {
            retryBtn.addEventListener('click', function() {
                initializeCanvasApp();
            });
        }
    });

    console.log("MessageRead.js loaded with extraction capabilities (Vanilla JS)");

})();
