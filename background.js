/* ---------- CHROME TAB CONTROL SYSTEM ---------- */

// Tab management functions
const TabController = {
  
  // Get all tabs
  getAllTabs: async function() {
    try {
      return await chrome.tabs.query({});
    } catch (error) {
      console.error('[TabController] Error getting tabs:', error);
      return [];
    }
  },
  
  // Get active tab
  getActiveTab: async function() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      return tab;
    } catch (error) {
      console.error('[TabController] Error getting active tab:', error);
      return null;
    }
  },
  
  // Create new tab
  createTab: async function(url, options = {}) {
    try {
      const tabOptions = {
        url: url,
        active: options.active !== false, // Default to active
        ...options
      };
      return await chrome.tabs.create(tabOptions);
    } catch (error) {
      console.error('[TabController] Error creating tab:', error);
      return null;
    }
  },
  
  // Close tab
  closeTab: async function(tabId) {
    try {
      if (Array.isArray(tabId)) {
        await chrome.tabs.remove(tabId);
      } else {
        await chrome.tabs.remove([tabId]);
      }
      return true;
    } catch (error) {
      console.error('[TabController] Error closing tab:', error);
      return false;
    }
  },
  
  // Switch to tab
  switchToTab: async function(tabId) {
    try {
      await chrome.tabs.update(tabId, { active: true });
      const tab = await chrome.tabs.get(tabId);
      await chrome.windows.update(tab.windowId, { focused: true });
      return true;
    } catch (error) {
      console.error('[TabController] Error switching to tab:', error);
      return false;
    }
  },
  
  // Duplicate tab
  duplicateTab: async function(tabId = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return null;
      
      return await chrome.tabs.duplicate(tabId);
    } catch (error) {
      console.error('[TabController] Error duplicating tab:', error);
      return null;
    }
  },
  
  // Pin/unpin tab
  togglePinTab: async function(tabId = null, pinned = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      const currentTab = await chrome.tabs.get(tabId);
      const shouldPin = pinned !== null ? pinned : !currentTab.pinned;
      
      await chrome.tabs.update(tabId, { pinned: shouldPin });
      return true;
    } catch (error) {
      console.error('[TabController] Error toggling pin:', error);
      return false;
    }
  },
  
  // Mute/unmute tab
  toggleMuteTab: async function(tabId = null, muted = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      const currentTab = await chrome.tabs.get(tabId);
      const shouldMute = muted !== null ? muted : !currentTab.mutedInfo.muted;
      
      await chrome.tabs.update(tabId, { muted: shouldMute });
      return true;
    } catch (error) {
      console.error('[TabController] Error toggling mute:', error);
      return false;
    }
  },
  
  // Move tab to position
  moveTab: async function(tabId, index) {
    try {
      await chrome.tabs.move(tabId, { index: index });
      return true;
    } catch (error) {
      console.error('[TabController] Error moving tab:', error);
      return false;
    }
  },
  
  // Group tabs
  groupTabs: async function(tabIds, groupOptions = {}) {
    try {
      const groupId = await chrome.tabs.group({ tabIds: tabIds });
      
      if (groupOptions.title || groupOptions.color) {
        await chrome.tabGroups.update(groupId, {
          title: groupOptions.title,
          color: groupOptions.color
        });
      }
      
      return groupId;
    } catch (error) {
      console.error('[TabController] Error grouping tabs:', error);
      return null;
    }
  },
  
  // Ungroup tabs
  ungroupTabs: async function(tabIds) {
    try {
      await chrome.tabs.ungroup(tabIds);
      return true;
    } catch (error) {
      console.error('[TabController] Error ungrouping tabs:', error);
      return false;
    }
  },
  
  // Close tabs to the right
  closeTabsToRight: async function(tabId = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      const currentTab = await chrome.tabs.get(tabId);
      const allTabs = await chrome.tabs.query({ windowId: currentTab.windowId });
      
      const tabsToClose = allTabs
        .filter(tab => tab.index > currentTab.index)
        .map(tab => tab.id);
      
      if (tabsToClose.length > 0) {
        await chrome.tabs.remove(tabsToClose);
      }
      
      return true;
    } catch (error) {
      console.error('[TabController] Error closing tabs to right:', error);
      return false;
    }
  },
  
  // Close other tabs
  closeOtherTabs: async function(tabId = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      const currentTab = await chrome.tabs.get(tabId);
      const allTabs = await chrome.tabs.query({ windowId: currentTab.windowId });
      
      const tabsToClose = allTabs
        .filter(tab => tab.id !== tabId && !tab.pinned)
        .map(tab => tab.id);
      
      if (tabsToClose.length > 0) {
        await chrome.tabs.remove(tabsToClose);
      }
      
      return true;
    } catch (error) {
      console.error('[TabController] Error closing other tabs:', error);
      return false;
    }
  },
  
  // Reload tab
  reloadTab: async function(tabId = null, bypassCache = false) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      await chrome.tabs.reload(tabId, { bypassCache: bypassCache });
      return true;
    } catch (error) {
      console.error('[TabController] Error reloading tab:', error);
      return false;
    }
  },
  
  // Go back in tab history
  goBack: async function(tabId = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      await chrome.tabs.goBack(tabId);
      return true;
    } catch (error) {
      console.error('[TabController] Error going back:', error);
      return false;
    }
  },
  
  // Go forward in tab history
  goForward: async function(tabId = null) {
    try {
      if (!tabId) {
        const activeTab = await this.getActiveTab();
        tabId = activeTab?.id;
      }
      if (!tabId) return false;
      
      await chrome.tabs.goForward(tabId);
      return true;
    } catch (error) {
      console.error('[TabController] Error going forward:', error);
      return false;
    }
  },
  
  // Search tabs by title or URL
  searchTabs: async function(query) {
    try {
      const allTabs = await this.getAllTabs();
      const searchQuery = query.toLowerCase();
      
      return allTabs.filter(tab => 
        tab.title.toLowerCase().includes(searchQuery) ||
        tab.url.toLowerCase().includes(searchQuery)
      );
    } catch (error) {
      console.error('[TabController] Error searching tabs:', error);
      return [];
    }
  }
};

const DIAG_LOG_KEY = 'holorunDiagnosticsLog';
const DIAG_LOG_LIMIT = 200;

function appendDiagLog(entry) {
  try {
    if (chrome.storage && chrome.storage.local) {
      chrome.storage.local.get([DIAG_LOG_KEY], (res) => {
        const logs = Array.isArray(res && res[DIAG_LOG_KEY]) ? res[DIAG_LOG_KEY] : [];
        logs.push(entry);
        if (logs.length > DIAG_LOG_LIMIT) {
          logs.splice(0, logs.length - DIAG_LOG_LIMIT);
        }
        chrome.storage.local.set({ [DIAG_LOG_KEY]: logs });
      });
    }
  } catch {
    // Ignore storage errors.
  }
}

function holoDiagLog(event, data) {
  const entry = {
    ts: Date.now(),
    source: 'background',
    event: event,
    data: data === undefined ? null : data
  };
  try {
    if (console && console.log) {
      if (data === undefined) {
        console.log(`[HoloDiag] ${event}`);
      } else {
        console.log(`[HoloDiag] ${event}`, data);
      }
    }
  } catch {
    // Ignore console errors.
  }
  appendDiagLog(entry);
}

const tabActivity = {};

function recordTabActivity(tabId, payload = {}) {
  if (!tabId) return;
  if (!tabActivity[tabId]) {
    tabActivity[tabId] = { clickCount: 0, lastClick: 0 };
  }
  if (payload.action === 'click') {
    tabActivity[tabId].clickCount += 1;
    tabActivity[tabId].lastClick = payload.timestamp || Date.now();
  }
}

/* ---------- MESSAGE HANDLERS FOR TAB CONTROL ---------- */
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type && message.type.startsWith('TAB_')) {
    handleTabMessage(message, sender, sendResponse);
    return true; // Keep channel open for async response
  }
});

async function handleTabMessage(message, sender, sendResponse) {
  try {
    let result;
    
    switch (message.type) {
      case 'TAB_GET_ALL':
        result = await TabController.getAllTabs();
        break;
        
      case 'TAB_GET_ACTIVE':
        result = await TabController.getActiveTab();
        break;
        
      case 'TAB_CREATE':
        result = await TabController.createTab(message.url, message.options);
        break;
        
      case 'TAB_CLOSE':
        result = await TabController.closeTab(message.tabId);
        break;
        
      case 'TAB_SWITCH':
        result = await TabController.switchToTab(message.tabId);
        break;
        
      case 'TAB_DUPLICATE':
        result = await TabController.duplicateTab(message.tabId);
        break;
        
      case 'TAB_TOGGLE_PIN':
        result = await TabController.togglePinTab(message.tabId, message.pinned);
        break;
        
      case 'TAB_TOGGLE_MUTE':
        result = await TabController.toggleMuteTab(message.tabId, message.muted);
        break;
        
      case 'TAB_MOVE':
        result = await TabController.moveTab(message.tabId, message.index);
        break;
        
      case 'TAB_GROUP':
        result = await TabController.groupTabs(message.tabIds, message.groupOptions);
        break;
        
      case 'TAB_UNGROUP':
        result = await TabController.ungroupTabs(message.tabIds);
        break;
        
      case 'TAB_CLOSE_TO_RIGHT':
        result = await TabController.closeTabsToRight(message.tabId);
        break;
        
      case 'TAB_CLOSE_OTHERS':
        result = await TabController.closeOtherTabs(message.tabId);
        break;
        
      case 'TAB_RELOAD':
        result = await TabController.reloadTab(message.tabId, message.bypassCache);
        break;
        
      case 'TAB_GO_BACK':
        result = await TabController.goBack(message.tabId);
        break;
        
      case 'TAB_GO_FORWARD':
        result = await TabController.goForward(message.tabId);
        break;
        
      case 'TAB_SEARCH':
        result = await TabController.searchTabs(message.query);
        break;
        
      default:
        result = { error: 'Unknown tab operation' };
    }

    try {
      const meta = { type: message.type };
      if (message.type === 'TAB_GET_ALL' && Array.isArray(result)) {
        meta.count = result.length;
      }
      if (message.type === 'TAB_SEARCH' && Array.isArray(result)) {
        meta.count = result.length;
      }
      if (message.type === 'TAB_SWITCH') {
        meta.tabId = message.tabId;
      }
      holoDiagLog('tab.op', meta);
    } catch {
      // Ignore diag log errors.
    }
    
    sendResponse({ success: true, result: result });
  } catch (error) {
    console.error('[TabController] Message handler error:', error);
    sendResponse({ success: false, error: error.message });
  }
}

function createContextMenu() {
  if (!chrome.contextMenus) return;

  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: "tag-text",
      title: "Tag",
      contexts: ["selection"]
    });
  });
}

/* ---------- INSTALL ---------- */
chrome.runtime.onInstalled.addListener(() => {
  createContextMenu();
});

/* ---------- STARTUP (CRITICAL) ---------- */
chrome.runtime.onStartup.addListener(() => {
  createContextMenu();
});

/* ---------- CLICK HANDLER ---------- */
chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== "tag-text") return;
  if (!tab?.id) return;

  chrome.tabs.sendMessage(tab.id, { type: "OPEN_TAG_POPUP" });
});

/* ---------- KEEP-ALIVE: Prevent service worker termination ---------- */
// Ping every 20 seconds to keep service worker alive during active usage
let keepAliveInterval = null;

function startKeepAlive() {
  if (keepAliveInterval) return;
  keepAliveInterval = setInterval(() => {
    chrome.runtime.getPlatformInfo(() => {
      // No-op; just keeps service worker alive
    });
  }, 20000);
}

function stopKeepAlive() {
  if (keepAliveInterval) {
    clearInterval(keepAliveInterval);
    keepAliveInterval = null;
  }
}

// Start keep-alive on install/startup
chrome.runtime.onInstalled.addListener(startKeepAlive);
chrome.runtime.onStartup.addListener(startKeepAlive);

// Also start when content script connects
chrome.runtime.onConnect.addListener(() => {
  startKeepAlive();
});

// Message handler to respond to pings from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.type === 'KEEP_ALIVE_PING') {
    sendResponse({ alive: true });
    return true;
  }
  
  if (request.type === 'GET_TAB_COUNT') {
    chrome.tabs.query({}, (tabs) => {
      sendResponse({ count: tabs.length });
    });
    return true; // Indicates async response
  }

  if (request.type === 'TAB_ACTIVITY') {
    const tabId = sender && sender.tab ? sender.tab.id : null;
    recordTabActivity(tabId, request || {});
    try {
      holoDiagLog('tab.activity', { tabId: tabId, action: request.action || null });
    } catch {
      // Ignore diag log errors.
    }
    sendResponse({ ok: true });
    return true;
  }

  if (request.type === 'TAB_ACTIVITY_SNAPSHOT') {
    try {
      holoDiagLog('tab.activity_snapshot', { count: Object.keys(tabActivity).length });
    } catch {
      // Ignore diag log errors.
    }
    sendResponse({ data: tabActivity });
    return true;
  }

  if (request.type === 'CAPTURE_TAB_SCREENSHOT') {
    const tabId = request.tabId;
    console.log('📸 Capturing screenshot for tab:', tabId);
    
    (async () => {
      try {
        // Get current active tab to restore later
        const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });
        const originalTabId = activeTab?.id;
        
        // Switch to the target tab
        await chrome.tabs.update(tabId, { active: true });
        
        // Small delay to let tab render
        await new Promise(resolve => setTimeout(resolve, 200));
        
        // Capture the now-visible tab
        const dataUrl = await chrome.tabs.captureVisibleTab(null, { format: 'jpeg', quality: 70 });
        
        // Switch back to original tab
        if (originalTabId && originalTabId !== tabId) {
          await chrome.tabs.update(originalTabId, { active: true });
        }
        
        console.log('✅ Screenshot captured for tab:', tabId);
        sendResponse({ success: true, dataUrl: dataUrl });
      } catch (error) {
        console.error('❌ Screenshot error for tab', tabId, ':', error);
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true; // Indicates async response
  }
