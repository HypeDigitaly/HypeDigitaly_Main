export const BrowserDataExtension = {
  name: "BrowserData",
  type: "effect",
  match: ({ trace }) => 
    trace.type === "ext_browserData" || 
    trace.payload?.name === "ext_browserData",
  effect: async ({ trace }) => {

    // Initialize error collector
    let errorMessages = [];

    // Get basic allowed cookies with error handling
    const getBasicCookies = () => {
      try {
        const allowedCookies = [
          'theme',          
          'language',       
          'display_mode',   
          'font_size'       
        ];
        
        return {
          success: true,
          data: document.cookie.split(';').reduce((acc, cookie) => {
            const [name, value] = cookie.split('=').map(c => c.trim());
            if (allowedCookies.includes(name)) {
              acc[name] = value;
            }
            return acc;
          }, {})
        };
      } catch (error) {
        errorMessages.push({
          function: 'getBasicCookies',
          error: error.message
        });
        return {
          success: false,
          data: {}
        };
      }
    };

    // Get URL related info with error handling
    const getUrlInfo = () => {
      try {
        return {
          success: true,
          data: {
            fullUrl: window.location.href,
            path: window.location.pathname,
            queryParams: new URLSearchParams(window.location.search).toString(),
            hash: window.location.hash
          }
        };
      } catch (error) {
        errorMessages.push({
          function: 'getUrlInfo',
          error: error.message
        });
        return {
          success: false,
          data: {
            fullUrl: 'unknown',
            path: 'unknown',
            queryParams: '',
            hash: ''
          }
        };
      }
    };

    // Get browser and hardware info with error handling
    const getSystemInfo = () => {
      try {
        const userAgent = navigator.userAgent;
        const systemInfo = {
          browser: {
            name: "Unknown",
            language: navigator.language || "Unknown",
            cookiesEnabled: navigator.cookieEnabled
          },
          device: {
            type: "desktop",
            platform: navigator.platform || "Unknown",
            os: "Unknown",
            processors: navigator.hardwareConcurrency || "Unknown",
            memory: navigator.deviceMemory || "Unknown", 
            touchPoints: navigator.maxTouchPoints || 0
          },
          display: {
            screen: {
              width: window.screen?.width || "Unknown",
              height: window.screen?.height || "Unknown",
              colorDepth: window.screen?.colorDepth || "Unknown",
              pixelRatio: window.devicePixelRatio || 1
            },
            viewport: {
              width: window.innerWidth || "Unknown",
              height: window.innerHeight || "Unknown"
            }
          }
        };

        try {
          // Detect browser name
          if (/chrome/i.test(userAgent)) {
            systemInfo.browser.name = "Chrome";
          } else if (/firefox/i.test(userAgent)) {
            systemInfo.browser.name = "Firefox"; 
          } else if (/safari/i.test(userAgent) && !/chrome/i.test(userAgent)) {
            systemInfo.browser.name = "Safari";
          } else if (/msie/i.test(userAgent) || /trident/i.test(userAgent)) {
            systemInfo.browser.name = "Internet Explorer";
          }
        } catch (browserError) {
          errorMessages.push({
            function: 'getSystemInfo - browser detection',
            error: browserError.message
          });
        }

        try {
          // Detect device type
          if (/mobile|android|iphone/i.test(userAgent)) {
            systemInfo.device.type = "mobile";
          } else if (/tablet|ipad/i.test(userAgent)) {
            systemInfo.device.type = "tablet";
          }
        } catch (deviceError) {
          errorMessages.push({
            function: 'getSystemInfo - device detection',
            error: deviceError.message
          });
        }

        try {
          // Detect operating system
          if (/windows/i.test(userAgent)) {
            systemInfo.device.os = "Windows";
          } else if (/macintosh|mac os x/i.test(userAgent)) {
            systemInfo.device.os = "MacOS"; 
          } else if (/linux/i.test(userAgent)) {
            systemInfo.device.os = "Linux";
          } else if (/android/i.test(userAgent)) {
            systemInfo.device.os = "Android";
          } else if (/ios/i.test(userAgent)) {
            systemInfo.device.os = "iOS";
          }
        } catch (osError) {
          errorMessages.push({
            function: 'getSystemInfo - OS detection',
            error: osError.message
          });
        }

        return {
          success: true,
          data: systemInfo
        };
      } catch (error) {
        errorMessages.push({
          function: 'getSystemInfo',
          error: error.message
        });
        return {
          success: false,
          data: {
            browser: { name: "Unknown", language: "Unknown", cookiesEnabled: false },
            device: { type: "Unknown", platform: "Unknown", os: "Unknown", processors: "Unknown", memory: "Unknown", touchPoints: 0 },
            display: { screen: { width: "Unknown", height: "Unknown", colorDepth: "Unknown", pixelRatio: 1 }, viewport: { width: "Unknown", height: "Unknown" } }
          }
        };
      }
    };

    // Get IP address with error handling
    const getIpAddress = async () => {
      try {
        const response = await fetch('https://api.ipify.org?format=json');
        const data = await response.json();
        return {
          success: true,
          data: data.ip
        };
      } catch (error) {
        errorMessages.push({
          function: 'getIpAddress',
          error: error.message
        });
        return {
          success: false,
          data: "Unknown"
        };
      }
    };

    // Collect all data with error handling
    let timestamp, timezone;
    try {
      timestamp = new Date().toISOString();
      timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    } catch (error) {
      errorMessages.push({
        function: 'getTimeInfo',
        error: error.message
      });
      timestamp = "Unknown";
      timezone = "Unknown";
    }

    // Collect all data
    const ipResult = await getIpAddress();
    const cookiesResult = getBasicCookies();
    const systemResult = getSystemInfo();
    const urlResult = getUrlInfo();

    // Send collected data with error information
    window.voiceflow.chat.interact({
      type: "complete",
      payload: {
        ip_address: ipResult.data,
        timestamp: timestamp,
        timezone: timezone,
        url: urlResult.data.fullUrl,
        path: urlResult.data.path,
        queryParams: urlResult.data.queryParams,
        urlHash: urlResult.data.hash,
        system: systemResult.data,
        basic_cookies: cookiesResult.data,
        error_info: {
          has_errors: errorMessages.length > 0,
          error_count: errorMessages.length,
          errors: errorMessages
        }
      }
    });
  }
};
