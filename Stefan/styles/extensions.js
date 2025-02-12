export const BrowserDataExtension = {
  name: "BrowserData",
  type: "effect",
  match: ({ trace }) => 
    trace.type === "ext_browserData" || 
    trace.payload?.name === "ext_browserData",
  effect: async ({ trace }) => {
    const getDeviceType = () => {
      const ua = navigator.userAgent;
      if (/(tablet|ipad|playbook|silk)|(android(?!.*mobi))/i.test(ua)) {
        return "tablet";
      }
      if (/Mobile|Android|iP(hone|od)|IEMobile|BlackBerry|Kindle|Silk-Accelerated|(hpw|web)OS|Opera M(obi|ini)/.test(ua)) {
        return "mobile";
      }
      return "desktop";
    };

    const getOS = () => {
      const userAgent = window.navigator.userAgent;
      const platform = window.navigator.platform;
      const macosPlatforms = ['Macintosh', 'MacIntel', 'MacPPC', 'Mac68K'];
      const windowsPlatforms = ['Win32', 'Win64', 'Windows', 'WinCE'];
      const iosPlatforms = ['iPhone', 'iPad', 'iPod'];

      if (macosPlatforms.indexOf(platform) !== -1) {
        return 'MacOS';
      } else if (iosPlatforms.indexOf(platform) !== -1) {
        return 'iOS';
      } else if (windowsPlatforms.indexOf(platform) !== -1) {
        return 'Windows';
      } else if (/Android/.test(userAgent)) {
        return 'Android';
      } else if (/Linux/.test(platform)) {
        return 'Linux';
      }
      return 'Unknown';
    };

    const getBrowserInfo = () => {
      const userAgent = navigator.userAgent.toLowerCase();
      let browserName = "Unknown";
      let browserVersion = "Unknown";
      let engineName = "Unknown";
      let engineVersion = "Unknown";

      // Detect browser and version
      if (userAgent.includes("firefox")) {
        browserName = "Firefox";
        browserVersion = userAgent.match(/firefox\/([\d.]+)/)?.[1] || "Unknown";
        engineName = "Gecko";
      } else if (userAgent.includes("chrome")) {
        browserName = "Chrome";
        browserVersion = userAgent.match(/chrome\/([\d.]+)/)?.[1] || "Unknown";
        engineName = "Blink";
      } else if (userAgent.includes("safari") && !userAgent.includes("chrome")) {
        browserName = "Safari";
        browserVersion = userAgent.match(/version\/([\d.]+)/)?.[1] || "Unknown";
        engineName = "WebKit";
      } else if (userAgent.includes("edg")) {
        browserName = "Edge";
        browserVersion = userAgent.match(/edg\/([\d.]+)/)?.[1] || "Unknown";
        engineName = "Blink";
      } else if (userAgent.includes("opera") || userAgent.includes("opr")) {
        browserName = "Opera";
        browserVersion = userAgent.match(/(?:opera|opr)\/([\d.]+)/)?.[1] || "Unknown";
        engineName = "Blink";
      }

      return { browserName, browserVersion, engineName, engineVersion };
    };

    const getDeviceVendor = () => {
      const ua = navigator.userAgent;
      if (/iPad|iPhone|iPod/.test(ua)) {
        return "Apple";
      } else if (/Android/.test(ua)) {
        const match = ua.match(/;\s([^;]+)\sBuild/);
        return match ? match[1].split(' ')[0] : "Unknown";
      }
      return "Unknown";
    };

    const getIpAddress = async () => {
      try {
        const response = await fetch('https://api.ipify.org?format=json');
        const data = await response.json();
        return data.ip;
      } catch (error) {
        return "Unable to fetch IP address";
      }
    };

    // Collect all system information using native browser APIs
    const { browserName, browserVersion, engineName, engineVersion } = getBrowserInfo();
    const deviceType = getDeviceType();
    const os = getOS();
    const ip = await getIpAddress();

    const systemInfo = {
      ip: ip,
      url: {
        current: window.location.href,
        pathname: window.location.pathname,
        search: window.location.search,
        params: Object.fromEntries(new URLSearchParams(window.location.search))
      },
      browser: {
        name: browserName,
        version: browserVersion,
        language: navigator.language,
        engine: engineName,
        engineVersion: engineVersion
      },
      device: {
        type: deviceType,
        platform: navigator.platform,
        os: os,
        osVersion: "Unknown", // Hard to reliably detect without ClientJS
        vendor: getDeviceVendor(),
        processors: navigator.hardwareConcurrency || 'Unknown',
        memory: navigator.deviceMemory ? `${navigator.deviceMemory}GB` : 'Unknown',
        touchPoints: navigator.maxTouchPoints || 0
      },
      display: {
        screen: {
          width: window.screen.width,
          height: window.screen.height,
          colorDepth: window.screen.colorDepth,
          pixelRatio: window.devicePixelRatio
        },
        availableResolution: {
          width: window.screen.availWidth,
          height: window.screen.availHeight
        }
      },
      system: {
        timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        systemLanguage: navigator.language,
        userAgent: navigator.userAgent
      }
    };

    // Send the collected data
    window.voiceflow.chat.interact({
      type: "complete",
      payload: systemInfo
    });
  }
};
