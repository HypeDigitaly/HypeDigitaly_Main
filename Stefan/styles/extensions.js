export const BrowserDataExtension = {
  name: "BrowserData",
  type: "effect",
  match: ({ trace }) => 
    trace.type === "ext_browserData" || 
    trace.payload?.name === "ext_browserData",
  effect: async ({ trace }) => {

    // Get basic allowed cookies only
    const getBasicCookies = () => {
      const allowedCookies = [
        'theme',          
        'language',       
        'display_mode',   
        'font_size'       
      ];
      
      return document.cookie.split(';').reduce((acc, cookie) => {
        const [name, value] = cookie.split('=').map(c => c.trim());
        if (allowedCookies.includes(name)) {
          acc[name] = value;
        }
        return acc;
      }, {});
    };

    // Get URL related info
    const getUrlInfo = () => {
      return {
        fullUrl: window.location.href,
        path: window.location.pathname,
        queryParams: new URLSearchParams(window.location.search).toString(),
        hash: window.location.hash
      };
    };

    // Get browser and hardware info
    const getSystemInfo = () => {
      const userAgent = navigator.userAgent;
      const systemInfo = {
        browser: {
          name: "Unknown",
          language: navigator.language,
          cookiesEnabled: navigator.cookieEnabled
        },
        device: {
          type: "desktop",
          platform: navigator.platform,
          os: "Unknown",
          processors: navigator.hardwareConcurrency || "Unknown",
          memory: navigator.deviceMemory || "Unknown", 
          touchPoints: navigator.maxTouchPoints || 0
        },
        display: {
          screen: {
            width: window.screen.width,
            height: window.screen.height,
            colorDepth: window.screen.colorDepth,
            pixelRatio: window.devicePixelRatio || 1
          },
          viewport: {
            width: window.innerWidth,
            height: window.innerHeight
          }
        }
      };

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

      // Detect device type
      if (/mobile|android|iphone/i.test(userAgent)) {
        systemInfo.device.type = "mobile";
      } else if (/tablet|ipad/i.test(userAgent)) {
        systemInfo.device.type = "tablet";
      }

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

      return systemInfo;
    };

    // Get IP address
    const getIpAddress = async () => {
      try {
        const response = await fetch('https://api.ipify.org?format=json');
        const data = await response.json();
        return data.ip;
      } catch (error) {
        return "Unable to fetch IP address";
      }
    };

    // Collect all data
    const ip = await getIpAddress();
    const cookies = getBasicCookies();
    const systemInfo = getSystemInfo();
    const urlInfo = getUrlInfo();
    const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const timestamp = new Date().toISOString();

    // Send collected data
    window.voiceflow.chat.interact({
      type: "complete",
      payload: {
        ip_address: ip,
        timestamp: timestamp,
        timezone: timezone,
        url: urlInfo.fullUrl,
        path: urlInfo.path,
        queryParams: urlInfo.queryParams,
        urlHash: urlInfo.hash,
        system: {
          browser: systemInfo.browser,
          device: systemInfo.device,
          display: systemInfo.display
        },
        basic_cookies: cookies
      }
    });
  }
};
