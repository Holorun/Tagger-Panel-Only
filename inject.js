// --- Holorun Tagger: SPA + DOM observers for reliable restoration ---
const DIAG_LOG_KEY = 'holorunDiagnosticsLog';
const DIAG_LOG_LIMIT = 200;

function appendDiagLog(entry) {
  try {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      chrome.storage.local.get([DIAG_LOG_KEY], (res) => {
        const logs = Array.isArray(res && res[DIAG_LOG_KEY]) ? res[DIAG_LOG_KEY] : [];
        logs.push(entry);
        if (logs.length > DIAG_LOG_LIMIT) {
          logs.splice(0, logs.length - DIAG_LOG_LIMIT);
        }
        chrome.storage.local.set({ [DIAG_LOG_KEY]: logs });
      });
      return;
    }
  } catch {
    // Ignore and fallback.
  }

  try {
    const raw = localStorage.getItem(DIAG_LOG_KEY);
    let logs = [];
    try {
      logs = raw ? JSON.parse(raw) : [];
    } catch {
      logs = [];
    }
    if (!Array.isArray(logs)) logs = [];
    logs.push(entry);
    if (logs.length > DIAG_LOG_LIMIT) {
      logs.splice(0, logs.length - DIAG_LOG_LIMIT);
    }
    localStorage.setItem(DIAG_LOG_KEY, JSON.stringify(logs));
  } catch {
    // Ignore storage errors.
  }
}

function holoDiagLog(event, data) {
  const entry = {
    ts: Date.now(),
    source: 'content',
    event: event,
    data: data === undefined ? null : data
  };
  try {
    if (typeof console !== 'undefined' && console.log) {
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

if (!window.holorunDiagnostics) {
  window.holorunDiagnostics = {
    getLogs: function(callback) {
      try {
        if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
          chrome.storage.local.get([DIAG_LOG_KEY], (res) => {
            const logs = Array.isArray(res && res[DIAG_LOG_KEY]) ? res[DIAG_LOG_KEY] : [];
            if (typeof callback === 'function') callback(logs);
          });
          return;
        }
      } catch {
        // Ignore and fallback.
      }

      try {
        const raw = localStorage.getItem(DIAG_LOG_KEY);
        const logs = raw ? JSON.parse(raw) : [];
        if (typeof callback === 'function') callback(Array.isArray(logs) ? logs : []);
      } catch {
        if (typeof callback === 'function') callback([]);
      }
    },
    clearLogs: function(callback) {
      try {
        if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
          chrome.storage.local.set({ [DIAG_LOG_KEY]: [] }, () => {
            if (typeof callback === 'function') callback();
          });
          return;
        }
      } catch {
        // Ignore and fallback.
      }

      try {
        localStorage.setItem(DIAG_LOG_KEY, JSON.stringify([]));
      } catch {
        // Ignore storage errors.
      }
      if (typeof callback === 'function') callback();
    }
  };
}
// These hooks ensure restoration re-runs when the site navigates between conversations
// or when dynamic content finishes loading after initial hydration.
(() => {
  const logPrefix = '[HolorunTagger/Observers]';
  let navTimer = null;
  let mutTimer = null;
  let mutationObserver = null;
  let isCreatingHighlight = false;
  let isRestoring = false;
  let lastRestoreTime = 0;
  let lastJumpedTagId = null; // Track the last jumped highlight to prevent re-searching
  const MIN_RESTORE_INTERVAL = 5000; // Minimum 5 seconds between restorations (increased from 3s)
  const GMAIL_MIN_RESTORE_INTERVAL = 8000; // Gmail needs longer intervals due to dynamic content
  const JUMP_STYLE_DURATION_MS = 4000; // Keep jump highlight visible a bit longer

  function getElement(node) {
    return node && node.nodeType === 1 ? node : node && node.parentElement ? node.parentElement : null;
  }

  function isOurUI(node) {
    try {
      const el = getElement(node);
      if (!el || typeof el.closest !== 'function') return false;
      return Boolean(el.closest('#tag-panel') || el.closest('.holorun-popup'));
    } catch {
      return false;
    }
  }

  function contentReady() {
    // Check: enough text content and presence of main container
    try {
      const host = location.hostname || '';
      const isChatGPT = /chat\.openai\.com$/.test(host);
      const isGmail = /mail\.google\.com$/.test(host);
      const isClaude = /claude\.ai$/.test(host);
      
      // Site-specific content requirements
      let minContent;
      if (isChatGPT) {
        minContent = 15000; // ChatGPT needs more content for conversations
      } else if (isGmail) {
        minContent = 1000; // Gmail has less dynamic content, lower threshold
      } else if (isClaude) {
        minContent = 8000; // Claude similar to ChatGPT but less content
      } else {
        minContent = 2000; // Default for other sites
      }
      
      const textLen = (document.body && document.body.innerText || '').trim().length;
      const hasMain = !!document.querySelector('[role="main"], main, .chat, .conversation');
      
      // Gmail-specific: check for Gmail-specific elements indicating ready state
      if (isGmail) {
        const gmailLoaded = !!document.querySelector('.aeJ, [data-thread-id], .aDP, .nH');
        return textLen > minContent && gmailLoaded;
      }
      
      // Check that we have actual conversation content, not just sidebar
      // Look for message-like elements or substantial text outside navigation
      const mainContent = document.querySelector('[role="main"]');
      if (mainContent) {
        const mainText = (mainContent.innerText || '').trim().length;
        const hasContent = mainText > minContent; // Wait for substantial content (was 300)
        return hasContent;
      }
      
      // Fallback: check for substantial text and elements suggesting loaded state
      return textLen > minContent && hasMain;
    } catch {
      return false;
    }
  }

  function waitForContentReady(maxWait = 8000, interval = 400) {
    return new Promise((resolve, reject) => {
      const start = Date.now();
      const tick = () => {
        if (contentReady()) return resolve();
        if (Date.now() - start >= maxWait) return reject(new Error('Content not ready'));
        setTimeout(tick, interval);
      };
      tick();
    });
  }

  function cleanUrlParamOnce(delay = 1500) {
    try {
      if (!location.href.includes('holorunTagId')) return;
      setTimeout(() => {
        try {
          if (typeof cleanUrl === 'function') {
            const newUrl = cleanUrl(location.href);
            history.replaceState(null, '', newUrl);
            console.log(`${logPrefix} Cleaned holorunTagId from URL`);
          }
        } catch (e) {
          console.warn(`${logPrefix} URL clean error:`, e);
        }
      }, delay);
    } catch (e) {
      console.warn(`${logPrefix} cleanUrlParamOnce error:`, e);
    }
  }

  function scheduleRestore(reason = 'unknown', delay = 800) {
    try {
      // Skip if already restoring or too soon after last restore
      const now = Date.now();
      const host = location.hostname || '';
      const isGmail = /mail\.google\.com$/.test(host);
      const cooldownInterval = isGmail ? GMAIL_MIN_RESTORE_INTERVAL : MIN_RESTORE_INTERVAL;
      
      if (isRestoring) {
        console.log(`${logPrefix} Skipping restore (${reason}) - already restoring`);
        return;
      }
      if (now - lastRestoreTime < cooldownInterval) {
        console.log(`${logPrefix} Skipping restore (${reason}) - cooldown active (${Math.round((cooldownInterval - (now - lastRestoreTime)) / 1000)}s remaining)`);
        return;
      }
      
      if (navTimer) clearTimeout(navTimer);
      try {
        if (typeof holoDiagLog === 'function') {
          holoDiagLog('restore.schedule', { reason: reason, delay: delay });
        }
      } catch {
        // Ignore diag log errors.
      }
      navTimer = setTimeout(() => {
        try {
          console.log(`${logPrefix} Triggering restore due to: ${reason}`);
          const isChatGPT = /chat\.openai\.com$/.test(location.hostname || '');
          const maxWaitMs = isChatGPT ? 15000 : 7000;
          const intervalMs = isChatGPT ? 500 : 350;
          waitForContentReady(maxWaitMs, intervalMs) // Domain-aware wait: shorter for non-ChatGPT pages
            .then(() => {
              if (typeof restoreTags === 'function') restoreTags();
              if (typeof autoJumpIfTagIdPresent === 'function') {
                autoJumpIfTagIdPresent();
                cleanUrlParamOnce(1500);
              }
            })
            .catch(() => {
              // Even if not ready, still attempt a light restore to catch simple pages
              if (typeof restoreTags === 'function') restoreTags();
              if (typeof autoJumpIfTagIdPresent === 'function') {
                autoJumpIfTagIdPresent();
                cleanUrlParamOnce(2000);
              }
            });
        } catch (e) {
          console.warn(`${logPrefix} Restore execution error:`, e);
        }
      }, delay);
    } catch (e) {
      console.warn(`${logPrefix} scheduleRestore error:`, e);
    }
  }

  function setupSpaNavigationHooks() {
    try {
      if (window.__holorunSpaHooksInstalled) return;
      window.__holorunSpaHooksInstalled = true;

      const origPushState = history.pushState;
      const origReplaceState = history.replaceState;

      history.pushState = function (...args) {
        const res = origPushState.apply(this, args);
        scheduleRestore('history.pushState');
        return res;
      };

      history.replaceState = function (...args) {
        const res = origReplaceState.apply(this, args);
        scheduleRestore('history.replaceState');
        return res;
      };

      window.addEventListener('popstate', () => scheduleRestore('popstate'));
      // Handle BFCache restores and tab visibility/focus changes common in SPAs
      window.addEventListener('pageshow', (e) => {
        const reason = e && e.persisted ? 'pageshow(bfcache)' : 'pageshow';
        scheduleRestore(reason, 600);
      });
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') {
          scheduleRestore('visibilitychange:visible', 600);
        }
      });
      window.addEventListener('focus', () => scheduleRestore('window.focus', 600));
      console.log(`${logPrefix} SPA navigation hooks installed`);
    } catch (e) {
      console.warn(`${logPrefix} setupSpaNavigationHooks error:`, e);
    }
  }

  function setupDomMutationRestore() {
    try {
      if (mutationObserver) return;
      mutationObserver = new MutationObserver((mutations) => {
        // Ignore if we're currently creating highlights or restoring
        if (isCreatingHighlight || isRestoring) return;
        
        // Ignore changes originating from our own UI to prevent feedback loops
        const onlyOurUI = mutations.every(m => isOurUI(m.target));
        if (onlyOurUI) return;
        
        // Gmail-specific: ignore common Gmail UI mutations that don't affect content
        const host = location.hostname || '';
        const isGmail = /mail\.google\.com$/.test(host);
        if (isGmail) {
          const isGmailUI = mutations.every(m => {
            const target = m.target;
            if (!target || typeof target.closest !== 'function') return false;
            // Ignore Gmail toolbar, sidebar, and UI element changes
            return target.closest('.aeN, .aKz, .ar7, .aDP, .bhZ, .G-asx, [role="navigation"], [class*="toolbar"], [class*="sidebar"]');
          });
          if (isGmailUI) return;
        }
        
        // Check if mutations include our tag highlights being added/removed
        const hasHighlightMutations = mutations.some(m => {
          if (m.type === 'childList') {
            // Check added nodes
            const addedNodes = [...m.addedNodes];
            if (addedNodes.some(n => 
              n.classList && (
                n.classList.contains('tag-highlight') ||
                n.classList.contains('pending-highlight') ||
                n.classList.contains('tag-highlight-applied')
              )
            )) return true;
            
            // Check removed nodes
            const removedNodes = [...m.removedNodes];
            if (removedNodes.some(n => 
              n.classList && (
                n.classList.contains('tag-highlight') ||
                n.classList.contains('pending-highlight') ||
                n.classList.contains('tag-highlight-applied')
              )
            )) return true;
          }
          return false;
        });
        
        // Skip if we're currently creating highlights
        if (window.__holorunCreatingHighlight) {
          return;
        }
        
        // Skip if highlight elements were added/removed (our own restoration or React tearing them down)
        if (hasHighlightMutations) {
          console.log(`${logPrefix} Skipping restore - highlight mutations detected`);
          return;
        }
        
        // Debounce to avoid excessive restores during heavy DOM updates
        if (mutTimer) clearTimeout(mutTimer);
        const debounceDelay = isGmail ? 5000 : 2500; // Gmail gets longer debounce
        mutTimer = setTimeout(() => {
          scheduleRestore('MutationObserver');
        }, debounceDelay);
      });

      mutationObserver.observe(document.body, {
        childList: true,
        subtree: true,
      });
      console.log(`${logPrefix} DOM mutation observer installed`);
    } catch (e) {
      console.warn(`${logPrefix} setupDomMutationRestore error:`, e);
    }
  }

  // Initialize hooks immediately after script evaluation
  try {
    setupSpaNavigationHooks();
    setupDomMutationRestore();
  } catch (e) {
    console.warn(`${logPrefix} initialization error:`, e);
  }

  // Export flags for highlight creation and restoration
  window.__holorunCreatingHighlight = false;
  Object.defineProperty(window, '__holorunCreatingHighlight', {
    set: (val) => { isCreatingHighlight = val; },
    get: () => isCreatingHighlight
  });
  window.__holorunRestoring = false;
  Object.defineProperty(window, '__holorunRestoring', {
    set: (val) => { isRestoring = val; },
    get: () => isRestoring
  });

  // CACHING: Store highlight positions in memory, apply styling only on demand
  // This avoids React DOM conflicts by not keeping permanent DOM elements
  function cacheHighlightPosition(tag) {
    const matchText = getMatchableText(tag);
    const kind = tag.kind || classifyTagKind(matchText, tag.containerTag || '', tag.iconMeta);
    const bodyText = document.body.innerText;
    const bodyTextNorm = bodyText.replace(/\s+/g, ' ');
    const tagTextNorm = matchText.trim().replace(/\s+/g, ' ');
    
    // Find position in page text
    let position = bodyText.indexOf(matchText);
    if (position === -1) {
      position = bodyTextNorm.indexOf(tagTextNorm);
    }
    
    if (position >= 0) {
      highlightCache[tag.id] = {
        tag: tag,
        position: position,
        textLength: tag.text.length,
        cachedAt: Date.now()
      };
      console.log(`[Cache] Stored position for tag: ${tag.id.substring(0, 8)}`);
    }
  }
  
  // When user clicks panel → apply temporary visual styling to the cached text
  function applyHighlightStyling(tagId) {
    return new Promise((resolve) => {
      safeStorageGet({ tags: [] }, (res) => {
        const tags = res.tags || [];
        const tag = tags.find(t => t.id === tagId);
        if (!tag) {
          resolve();
          return;
        }

        const matchText = getMatchableText(tag);
      
        // If not cached, cache it now
        if (!highlightCache[tagId]) {
          cacheHighlightPosition(tag);
        }
        
        // FIRST: Try to find the existing highlight span in the DOM
        let existing = document.querySelector(`[data-tag-id="${tagId}"]`);
        if (existing) {
          console.log(`[StyleHighlight] ✓ Found existing highlight in DOM, scrolling to it`);

          // Diagnostic: Check element state before scroll
          const beforePos = existing.getBoundingClientRect();
          console.log(`[Scroll/Before] Element position: top=${beforePos.top}, left=${beforePos.left}, height=${beforePos.height}`);
          console.log(`[Scroll/Before] Window height: ${window.innerHeight}`);

          // Force display inline if not already
          existing.style.setProperty('display', 'inline', 'important');

          // Use robust scroll that accounts for sticky headers and inner containers
          ensureVisible(existing);

          // Verify and pulse after a short delay
          setTimeout(() => {
            const afterPos = existing.getBoundingClientRect();
            const inViewport = afterPos.top >= 0 && afterPos.top <= window.innerHeight;
            console.log(`[Scroll/After] Element position: top=${afterPos.top}, left=${afterPos.left}, height=${afterPos.height}`);
            console.log(`[Scroll/After] In viewport: ${inViewport}`);
            pulseHighlight(existing);
          }, 400);

          resolve();
          return;
        }
        
        // Icon tags: locate icon element and pulse its existing or new highlight
        if (tag.iconMeta) {
          const iconEl = findIconElement(tag.iconMeta);
          if (iconEl) {
            const highlight = wrapElementWithHighlight(iconEl, tag);
            if (highlight) {
              try { highlight.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' }); } catch (_) {}
              pulseHighlight(highlight);
              console.log(`[StyleHighlight] ✓ Pulsed icon tag: ${tagId.substring(0, 8)}`);
              resolve();
              return;
            }
          }
        }

        // FALLBACK: Find the actual text node and apply styling with anchor validation
        let range = findRangeByRawText(matchText, tag.context);
        if (!range) {
          range = findRangeByTokens(matchText, tagId);
        }
        if (!range) {
          range = findRangeByRawTextLoose(matchText);
        }
        if (!range) {
          console.warn(`[StyleHighlight] Could not find text for tag: ${tagId}`);
          resolve();
          return;
        }

        // CRITICAL: Validate range size before applying temporary highlight
        let rangeText = range.toString();
        const MAX_REASONABLE_HIGHLIGHT = 500; // chars
        
        // Clamp to a single block to avoid wrapping entire messages
        const clampedForJump = clampRangeToSingleBlock(range, MAX_REASONABLE_HIGHLIGHT);
        if (clampedForJump) {
          range = clampedForJump;
          rangeText = range.toString();
        }
        
        if (rangeText.length > MAX_REASONABLE_HIGHLIGHT) {
          console.warn(`[Jump/Validation] ❌ REJECTED: Range too large (${rangeText.length} chars, max ${MAX_REASONABLE_HIGHLIGHT})`);
          console.warn(`[Jump/Validation] Trimming range and creating highlight...`);
          const trimmed = trimRangeToLength(range, MAX_REASONABLE_HIGHLIGHT);
          if (trimmed) {
            range = trimmed;
          } else {
            console.warn(`[Jump/Validation] Trimming failed, aborting jump`);
            resolve();
            return;
          }
        }

        // Create temporary styled span
        window.__holorunCreatingHighlight = true; // guard MutationObserver during styling
        const span = document.createElement('span');
        span.className = 'tag-highlight-applied';
        span.id = `tag-styled-${tagId}`;
        // Minimal styling: single yellow background, no outlines or shadows
        span.style.backgroundColor = 'rgba(255, 255, 0, 0.35)';

        try {
          // Use cloneContents instead of extractContents to avoid React conflicts
          const contents = range.cloneContents();
          span.appendChild(contents);
          range.insertNode(span);
          
          // Scroll into view
          span.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
          
          // Remove styling after a short delay (visible pulse linger)
          setTimeout(() => {
            try {
              if (span.parentNode) {
                const parent = span.parentNode;
                while (span.firstChild) {
                  parent.insertBefore(span.firstChild, span);
                }
                parent.removeChild(span);
                parent.normalize();
                console.log(`[StyleHighlight] ✓ Styling removed after pulse for: ${tagId.substring(0, 8)}`);
              }
            } catch (e) {
              console.warn('[StyleHighlight] Cleanup error (React may have already removed node):', e.message);
            }
            window.__holorunCreatingHighlight = false; // clear after removal
          }, JUMP_STYLE_DURATION_MS);
          
          console.log(`[StyleHighlight] ✓ Applied temporary styling to: ${tagId.substring(0, 8)}`);
          resolve();
        } catch (e) {
          console.warn('[StyleHighlight] Failed to apply styling:', e.message);
          window.__holorunCreatingHighlight = false;
          resolve();
        }
      });
    });
  }
  
  // Make apply function globally accessible
  window.__applyHighlightStyling = applyHighlightStyling;
  // Expose caching helper so restoreTags can use it outside this closure
  window.cacheHighlightPosition = cacheHighlightPosition;
})();

// inject.js - FINAL FIXED VERSION

console.log("🚀 Holorun Tagger loaded on:", location.href);

// Helper: Normalize text by removing combining characters, smart quotes, dashes, and collapsing whitespace
// Uses NFD decomposition to handle all emoji/diacritic variations
// Examples: ✔️→✔, 1⃣→1, é→e (decomposed then combining marks removed)
// Also normalizes smart quotes, dashes, and other special punctuation
function normalizeTextForMatching(text) {
  if (!text) return '';
  return text
    .normalize('NFD')                 // Decompose (✔️ → ✔ + variation selector, etc.)
    .replace(/[\u0300-\u036F\u1AB0-\u1AFF\u1DC0-\u1DFF\u20D0-\u20FF\uFE00-\uFE0F]/g, '') // Remove all combining marks
    // Normalize smart punctuation to ASCII equivalents
    .replace(/[\u2018\u2019\u201C\u201D]/g, '"')  // Smart quotes → straight quotes
    .replace(/[\u2013\u2014]/g, '-')              // En dash, em dash → hyphen
    .replace(/[\u2026]/g, '...')                  // Ellipsis … → ...
    .replace(/[\u00A0\u2000-\u200B\u3000]/g, ' ') // Various spaces → regular space
    .replace(/\s+/g, ' ')             // Collapse all whitespace to single spaces
    .trim();                          // Remove leading/trailing whitespace
}

// Helper: remove holorunTagId params and trailing ?/# from a URL string
function cleanUrl(url) {
  return url
    .replace(/#holorunTagId=[^&]*/g, '')
    .replace(/\?holorunTagId=[^&]*/g, '')
    .replace(/#$/, '')
    .replace(/\?$/, '');
}

// Helper: Choose the best content root to avoid sidebars/nav
// Expand hidden/collapsed CONTENT sections (not UI controls) on any site
function expandCollapsedSections() {
  try {
    // GLOBAL: Only expand on ChatGPT/Claude - skip everything else
    const host = location.hostname || '';
    const isTargetSite = /chatgpt\.com$|chat\.openai\.com$|claude\.ai$/.test(host);
    
    if (!isTargetSite) {
      // Completely disabled on non-target sites
      return;
    }

    // Only search in main content area to avoid nav/UI
    const contentRoot = getContentRoot();
    if (!contentRoot) return;
    
    // Find collapsed sections, but exclude UI controls
    const expandables = contentRoot.querySelectorAll('[aria-expanded="false"]');
    
    console.log(`[ExpandContent] Found ${expandables.length} collapsed elements, filtering UI controls...`);
    
    let expanded = 0;
    expandables.forEach((el) => {
      // Skip if it's clearly a UI control
      const tag = el.tagName.toLowerCase();
      const ariaLabel = (el.getAttribute('aria-label') || '').toLowerCase();
      const role = el.getAttribute('role') || '';
      const classes = (el.className || '').toLowerCase();
      
      // Exclude UI control patterns
      const isUIControl = (
        tag === 'button' || 
        tag === 'a' ||
        role === 'menuitem' || 
        role === 'tab' || 
        role === 'combobox' ||
        role === 'switch' ||
        classes.includes('menu') ||
        classes.includes('dropdown') ||
        classes.includes('toggle') ||
        ariaLabel.includes('menu') ||
        ariaLabel.includes('dropdown') ||
        ariaLabel.includes('navigate')
      );
      
      if (isUIControl) {
        console.log(`[ExpandContent] ⊘ Skipping UI control:`, tag, role, classes.substring(0, 30));
        return;
      }
      
      try {
        // Likely a content section - expand it
        el.setAttribute('aria-expanded', 'true');
        el.click?.();
        expanded++;
      } catch (e) {
        // Silently fail for individual elements
      }
    });
    
    console.log(`[ExpandContent] ✓ Expanded ${expanded} content sections`);
  } catch (e) {
    console.warn('[ExpandContent] Error:', e.message);
  }
}

function getContentRoot() {
  try {
    const host = location.hostname || '';
    // Prefer explicit main area when present
    const prefer = (
      document.querySelector('[role="main"]') ||
      document.querySelector('main') ||
      document.querySelector('article')
    );

    // Site-specific tweaks
    if (/chatgpt\.com$|chat\.openai\.com$/.test(host)) {
      // ChatGPT: conversation area lives under role=main, sidebar is separate
      return prefer || document.body;
    }
    if (/claude\.ai$/.test(host)) {
      // Claude: main chat area is under main; avoid left nav
      return prefer || document.body;
    }
    if (/mail\.google\.com$/.test(host)) {
      // Gmail: focus on main content area, avoid sidebars and navigation
      const gmailMain = document.querySelector('.nH .aDF, .nH [role="main"], .aKz');
      return gmailMain || prefer || document.body;
    }
    if (/learn\.microsoft\.com$/.test(host)) {
      // MS Learn has a main element
      return prefer || document.body;
    }
    // Default
    return prefer || document.body;
  } catch (_) {
    return document.body;
  }
}

// Helper: Match URLs by origin + pathname (ignores query params, hash, etc.)
// This is the simpler and more reliable approach used by ChatGPT Highlighter
function samePageUrl(url1, url2) {
  try {
    const a = new URL(url1);
    const b = new URL(url2);
    return a.origin === b.origin && a.pathname === b.pathname;
  } catch {
    return url1 === url2;
  }
}

// Helper: Get hash parameter (for deep-linking support)
function getHashParam(key) {
  try {
    const hash = (location.hash || '').replace(/^#/, '');
    const params = new URLSearchParams(hash);
    return params.get(key);
  } catch {
    return null;
  }
}

// Helper: Check if extension context is still valid
function isExtensionContextValid() {
  try {
    // Try to access chrome object - this is the most reliable check
    // Don't check storage.local specifically as it may be temporarily unavailable
    // even when the extension context is otherwise valid
    return !!(chrome && chrome.runtime && chrome.runtime.sendMessage);
  } catch {
    return false;
  }
}

// Helper: Retry wrapper for chrome.storage calls with fallback
function safeStorageGet(keys, callback, retries = 3) {
  const attempt = (retriesLeft) => {
    try {
      // Try storage API directly - don't pre-check context as it can be flaky
      if (!chrome || !chrome.storage || !chrome.storage.local) {
        console.warn('[SafeStorage] chrome.storage.local unavailable, retries left:', retriesLeft);
        if (retriesLeft > 0) {
          setTimeout(() => attempt(retriesLeft - 1), 300);
          return;
        }
        callback({ tags: [] }); // Return empty on failure
        return;
      }
      
      chrome.storage.local.get(keys, (result) => {
        if (chrome.runtime.lastError) {
          console.warn('[SafeStorage] Get error:', chrome.runtime.lastError.message, 'retries left:', retriesLeft);
          if (retriesLeft > 0) {
            setTimeout(() => attempt(retriesLeft - 1), 300);
          } else {
            callback({ tags: [] });
          }
        } else {
          callback(result);
        }
      });
    } catch (e) {
      console.warn('[SafeStorage] Exception:', e.message, 'retries left:', retriesLeft);
      if (retriesLeft > 0) {
        setTimeout(() => attempt(retriesLeft - 1), 300);
      } else {
        callback({ tags: [] });
      }
    }
  };
  attempt(retries);
}

function safeStorageSet(items, callback, retries = 3) {
  const attempt = (retriesLeft) => {
    try {
      // Try storage API directly - don't pre-check context as it can be flaky
      if (!chrome || !chrome.storage || !chrome.storage.local) {
        console.warn('[SafeStorage] chrome.storage.local unavailable for set, retries left:', retriesLeft);
        if (retriesLeft > 0) {
          setTimeout(() => attempt(retriesLeft - 1), 300);
          return;
        }
        if (callback) callback();
        return;
      }
      
      chrome.storage.local.set(items, () => {
        if (chrome.runtime.lastError) {
          console.warn('[SafeStorage] Set error:', chrome.runtime.lastError.message, 'retries left:', retriesLeft);
          if (retriesLeft > 0) {
            setTimeout(() => attempt(retriesLeft - 1), 300);
          } else {
            if (callback) callback();
          }
        } else {
          if (callback) callback();
        }
      });
    } catch (e) {
      console.warn('[SafeStorage] Set exception:', e.message, 'retries left:', retriesLeft);
      if (retriesLeft > 0) {
        setTimeout(() => attempt(retriesLeft - 1), 300);
      } else {
        if (callback) callback();
      }
    }
  };
  attempt(retries);
}

// Periodic keep-alive ping to service worker
function pingServiceWorker() {
  try {
    if (chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage({ type: 'KEEP_ALIVE_PING' }, (response) => {
        if (chrome.runtime.lastError) {
          console.warn('[KeepAlive] Service worker not responding:', chrome.runtime.lastError.message);
        }
      });
    }
  } catch (e) {
    console.warn('[KeepAlive] Ping failed:', e.message);
  }
}

// Ping every 15 seconds to keep service worker alive during active sessions
setInterval(pingServiceWorker, 15000);

// State
let lastSelectionRange = null;
let pendingHighlightSpan = null;
let panel = null;
let scrollRestoreTimer = null;
let fullscreenOverlay = null;
const PANEL_ONLY_MODE = true;
function showFullscreenOverlay() {
  if (PANEL_ONLY_MODE) {
    console.log('Panel-only build: fullscreen mode disabled');
    return;
  }

  const panelHost = document.getElementById('tag-panel-host');
  const originalPageContent = document.body.innerHTML;
  
  if (panelHost) {
    panelHost.style.display = 'none';
  }
  
  if (document.fullscreenEnabled && document.body.requestFullscreen) {
    document.body.requestFullscreen().catch(() => {});
  }
  
  if (fullscreenOverlay) return;
  
  fullscreenOverlay = document.createElement('div');
  fullscreenOverlay.id = 'holorun-fullscreen-overlay';
  fullscreenOverlay.style.position = 'fixed';
  fullscreenOverlay.style.top = '0';
  fullscreenOverlay.style.left = '0';
  fullscreenOverlay.style.width = '100vw';
  fullscreenOverlay.style.height = '100vh';
  fullscreenOverlay.style.background = '#0b1b1f';
  fullscreenOverlay.style.zIndex = '2147483647';
  fullscreenOverlay.style.display = 'flex';
  fullscreenOverlay.style.flexDirection = 'column';
  fullscreenOverlay.style.justifyContent = 'flex-start';
  fullscreenOverlay.style.alignItems = 'center';
  fullscreenOverlay.style.transition = 'opacity 0.3s';
  fullscreenOverlay.style.opacity = '1';

  const exitFullscreenOverlay = () => {
    console.log('🚪 Exiting fullscreen...');
    document.removeEventListener('keydown', onKeyDown);
    fullscreenOverlay.style.opacity = '0';
    
    setTimeout(() => {
      console.log('✋ Restoring panel and page...');
      
      // Exit fullscreen mode first
      if (document.fullscreenElement) {
        document.exitFullscreen().catch(() => {});
      }
      
      // Remove the overlay
      if (fullscreenOverlay && fullscreenOverlay.parentNode) {
        fullscreenOverlay.remove();
      }
      fullscreenOverlay = null;
      
      // Restore the panel
      if (panelHost) {
        panelHost.style.display = 'block';
        panelHost.style.visibility = 'visible';
        panelHost.style.pointerEvents = 'auto';
        panelHost.style.zIndex = '2147483646';
      }
      
      console.log('✅ Fullscreen exit complete');
    }, 250);
  };

  const onKeyDown = (event) => {
    if (event.key === 'Escape') {
      exitFullscreenOverlay();
    }
  };
  document.addEventListener('keydown', onKeyDown);

  fullscreenOverlay.addEventListener('click', (event) => {
    if (event.target === fullscreenOverlay) {
      exitFullscreenOverlay();
    }
  });

  const layout = document.createElement('div');
  layout.id = 'holorun-fullscreen-organizer';
  layout.style.cssText = 'width: min(92vw, 1400px); margin-top: 26px; display: flex; flex-direction: column; gap: 16px;';

  const mainRow = document.createElement('div');
  mainRow.style.cssText = 'display: grid; grid-template-columns: 3fr 1.15fr; gap: 18px; align-items: start;';

  const heroEl = document.createElement('div');
  const sideEl = document.createElement('div');
  const miniEl = document.createElement('div');

  heroEl.id = 'fullscreen-workflow-hero';
  sideEl.id = 'fullscreen-workflow-side';
  miniEl.id = 'fullscreen-workflow-mini';

  heroEl.style.cssText = 'display: flex; flex-direction: column; min-height: 50vh; width: 100%;';
  sideEl.style.cssText = 'display: flex; flex-direction: column; gap: 14px;';
  miniEl.style.cssText = 'display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 14px;';

  mainRow.appendChild(heroEl);
  mainRow.appendChild(sideEl);

  layout.appendChild(mainRow);
  layout.appendChild(miniEl);
  fullscreenOverlay.appendChild(layout);

  const RELATIONS_KEY = 'holorunHoloRelations';
  const getRelations = (callback) => {
    try {
      if (chrome && chrome.storage && chrome.storage.local) {
        chrome.storage.local.get([RELATIONS_KEY], (res) => {
          callback(res && res[RELATIONS_KEY] ? res[RELATIONS_KEY] : { mainId: null, relatedMap: {} });
        });
        return;
      }
    } catch {
      // Ignore and fallback.
    }
    try {
      const raw = localStorage.getItem(RELATIONS_KEY);
      callback(raw ? JSON.parse(raw) : { mainId: null, relatedMap: {} });
    } catch {
      callback({ mainId: null, relatedMap: {} });
    }
  };

  const saveRelations = (data, callback) => {
    try {
      if (chrome && chrome.storage && chrome.storage.local) {
        chrome.storage.local.set({ [RELATIONS_KEY]: data }, () => callback && callback());
        return;
      }
    } catch {
      // Ignore and fallback.
    }
    try {
      localStorage.setItem(RELATIONS_KEY, JSON.stringify(data));
    } catch {
      // Ignore storage errors.
    }
    if (callback) callback();
  };

  const contextMenu = document.createElement('div');
  contextMenu.id = 'holorun-holo-context-menu';
  contextMenu.style.cssText = 'position: fixed; z-index: 2147483647; background: #e2e5e7; color: #1f2a2e; border: 1px solid rgba(0,0,0,0.2); border-radius: 6px; padding: 6px; display: none; min-width: 160px; box-shadow: 0 6px 18px rgba(0,0,0,0.3); font-size: 12px;';
  const relateBtn = document.createElement('button');
  relateBtn.textContent = 'Relate to main holo';
  relateBtn.style.cssText = 'width: 100%; padding: 6px 8px; background: #f0f2f3; border: 1px solid rgba(0,0,0,0.2); border-radius: 4px; cursor: pointer; font-size: 12px; text-align: left;';
  contextMenu.appendChild(relateBtn);
  fullscreenOverlay.appendChild(contextMenu);

  let contextTargetId = null;
  const hideContextMenu = () => {
    contextMenu.style.display = 'none';
    contextTargetId = null;
  };
  fullscreenOverlay.addEventListener('click', () => hideContextMenu());

  if (!window.chromeTabController) {
    const empty = document.createElement('div');
    empty.style.cssText = 'color: rgba(255, 255, 255, 0.7); font-size: 14px; padding: 24px;';
    empty.textContent = 'Tab controls unavailable in this context.';
    heroEl.appendChild(empty);
    document.body.appendChild(fullscreenOverlay);
    return;
  }

  const createIframeCard = (tab, size = 'mini', isMain = false) => {
        const card = document.createElement('button');
        card.className = 'workflow-iframe-card' + (tab.active ? ' active' : '');
        card.type = 'button';
        card.title = `Click to switch to: ${tab.title}`;
        card.dataset.tabId = tab.id;
        const minHeight = size === 'hero' ? '50vh' : '28vh';
        card.style.cssText = `
          background: rgba(0, 0, 0, 0.05);
          border: 2px solid rgba(255, 255, 255, 0.65);
          border-radius: 4px;
          padding: 0;
          cursor: pointer;
          transition: all 0.2s ease;
          display: flex;
          flex-direction: column;
          overflow: hidden;
          text-align: left;
          width: 100%;
          height: ${minHeight};
        `;
        
        card.onmouseenter = () => {
          card.style.borderColor = 'rgba(255, 255, 255, 0.85)';
          card.style.background = 'rgba(255, 255, 255, 0.08)';
          card.style.transform = 'translateY(-1px)';
          card.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.35)';
        };
        card.onmouseleave = () => {
          card.style.borderColor = tab.active ? 'rgba(255, 255, 255, 0.95)' : 'rgba(255, 255, 255, 0.65)';
          card.style.background = tab.active ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.05)';
          card.style.transform = 'translateY(0)';
          card.style.boxShadow = tab.active ? 'inset 0 0 0 1px rgba(255, 255, 255, 0.45)' : 'none';
        };
        
        if (tab.active) {
          card.style.borderColor = 'rgba(255, 255, 255, 0.95)';
          card.style.background = 'rgba(255, 255, 255, 0.1)';
          card.style.boxShadow = 'inset 0 0 0 1px rgba(255, 255, 255, 0.45)';
        }
        
        const chromeBar = document.createElement('div');
        chromeBar.style.cssText = `display: flex; align-items: center; gap: 6px; padding: ${size === 'hero' ? '6px 8px' : '5px 6px'}; background: #e2e5e7; border-bottom: 1px solid rgba(0, 0, 0, 0.18); z-index: 10; position: relative;`;
        
        const favicon = document.createElement('img');
        favicon.src = tab.favIconUrl || `https://www.google.com/s2/favicons?domain=${(() => {
          try {
            return new URL(tab.url).hostname;
          } catch {
            return '';
          }
        })()}`;
        favicon.style.cssText = 'width: 11px; height: 11px; flex-shrink: 0;';
        favicon.onerror = () => { favicon.style.display = 'none'; };
        
        const titleText = document.createElement('div');
        titleText.textContent = tab.title || tab.url || 'Untitled tab';
        titleText.style.cssText = `flex: 0 0 auto; font-size: ${size === 'hero' ? '11px' : '9px'}; color: #273338; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 120px;`;

        const urlBar = document.createElement('div');
        try {
          urlBar.textContent = new URL(tab.url).hostname;
        } catch {
          urlBar.textContent = tab.url;
        }
        urlBar.style.cssText = `flex: 1; font-size: ${size === 'hero' ? '10px' : '8px'}; color: #555; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;`;
        
        chromeBar.appendChild(favicon);
        chromeBar.appendChild(titleText);
        chromeBar.appendChild(urlBar);
        
        const previewContainer = document.createElement('div');
        previewContainer.style.cssText = `
          flex: 1; 
          background: linear-gradient(135deg, #1a3a42 0%, #0f2935 100%) center/cover no-repeat;
          position: relative; 
          overflow: hidden;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 8px;
        `;
        
        const loadingSpinner = document.createElement('div');
        loadingSpinner.textContent = '⏳';
        loadingSpinner.style.cssText = 'font-size: 24px; opacity: 0.6;';
        previewContainer.appendChild(loadingSpinner);
        
        card.appendChild(chromeBar);
        card.appendChild(previewContainer);
        
        // Capture screenshot asynchronously
        setTimeout(() => {
          chrome.runtime.sendMessage({ type: 'CAPTURE_TAB_SCREENSHOT', tabId: tab.id }, (response) => {
            if (response && response.success && response.dataUrl) {
              console.log('📸 Got screenshot for tab:', tab.id);
              loadingSpinner.remove();
              previewContainer.style.backgroundImage = `url('${response.dataUrl}')`;
              previewContainer.style.backgroundSize = 'cover';
              previewContainer.style.backgroundPosition = 'center';
            } else {
              console.warn('❌ Screenshot failed for tab:', tab.id);
              loadingSpinner.textContent = isMain ? '★ MAIN' : '👆 CLICK';
              loadingSpinner.style.cssText = `
                font-size: ${size === 'hero' ? '16px' : '11px'};
                color: rgba(255, 255, 255, 0.5);
                font-weight: 600;
              `;
            }
          });
        }, 100);
        
        card.onclick = (e) => {
          e.preventDefault();
          console.log('👆 Clicking tab:', tab.id, tab.title);
          if (ChromeTabController && ChromeTabController.switchToTab) {
            ChromeTabController.switchToTab(tab.id);
          } else {
            console.error('⚠️ ChromeTabController not available');
          }
        };
        
        card.oncontextmenu = (event) => {
          event.preventDefault();
          if (isMain) return;
          contextTargetId = tab.id;
          contextMenu.style.left = `${event.clientX}px`;
          contextMenu.style.top = `${event.clientY}px`;
          contextMenu.style.display = 'block';
        };
        
        return card;
      };
      
  const renderGrid = (tabs, relations) => {
    heroEl.innerHTML = '';
    sideEl.innerHTML = '';
    miniEl.innerHTML = '';
    miniEl.style.display = 'none';
    
    if (!tabs.length) {
      const empty = document.createElement('div');
      empty.style.cssText = 'text-align: center; padding: 40px 20px; color: rgba(255, 255, 255, 0.6); font-size: 13px;';
      empty.textContent = 'No tabs available.';
      heroEl.appendChild(empty);
      return;
    }
    const activeTab = tabs.find(tab => tab.active) || tabs[0];
    const mainId = relations.mainId || activeTab.id;
    const mainTab = tabs.find(tab => tab.id === mainId) || activeTab;
    const relatedIds = (relations.relatedMap && relations.relatedMap[mainTab.id]) || [];
    const relatedSet = new Set(relatedIds);

    const otherHolos = tabs.filter(tab => tab.id !== mainTab.id && !relatedSet.has(tab.id));
    
    console.log('🎯 renderGrid - Main:', mainTab.title, 'Others:', otherHolos.length, 'Tabs total:', tabs.length);

    heroEl.appendChild(createIframeCard(mainTab, 'hero', true));
    otherHolos.forEach(tab => sideEl.appendChild(createIframeCard(tab, 'side', false)));
  };

  const refreshTabs = async () => {
    try {
      const tabs = await ChromeTabController.getAllTabs();
      console.log('🔄 Loaded tabs:', tabs.length, tabs);
      getRelations((relations) => {
        console.log('📍 Relations:', relations);
        const activeTab = tabs.find(tab => tab.active) || tabs[0];
        if (!relations.mainId && activeTab) {
          relations.mainId = activeTab.id;
          saveRelations(relations);
        }
        renderGrid(tabs, relations);
      });
    } catch (err) {
      console.error('❌ Tab loading error:', err);
      const errorMsg = document.createElement('div');
      errorMsg.style.cssText = 'color: rgba(255, 255, 255, 0.7); font-size: 13px; padding: 24px;';
      errorMsg.textContent = 'Error loading tabs: ' + err.message;
      heroEl.appendChild(errorMsg);
    }
  };

  relateBtn.onclick = () => {
    if (!contextTargetId) return;
    getRelations((relations) => {
      const mainId = relations.mainId;
      if (!mainId) return;
      const relatedMap = relations.relatedMap || {};
      const list = Array.isArray(relatedMap[mainId]) ? relatedMap[mainId] : [];
      if (!list.includes(contextTargetId)) {
        list.push(contextTargetId);
      }
      relatedMap[mainId] = list;
      saveRelations({ mainId, relatedMap }, () => {
        hideContextMenu();
        refreshTabs();
      });
    });
  };

  refreshTabs();

  document.body.appendChild(fullscreenOverlay);
}

// Example: Attach to FAB/panel click
document.addEventListener('click', function(e) {
  if (PANEL_ONLY_MODE) return;
  const panelHost = document.getElementById('tag-panel-host');
  if (panelHost && e.target === panelHost) {
    showFullscreenOverlay();
  }
});

// Cache: Store highlight metadata in memory (not in DOM) to avoid React conflicts
const highlightCache = {};  // { tagId: { tagInfo, textInfo, appliedElement } }

// --- Helper: Extract surrounding context from a text node ---
function extractContext(startNode, endNode, highlightedText, contextLength = 100) {
  try {
    // SIMPLIFIED: Always use exactly 30 chars (ChatGPT Highlighter approach)
    const host = (startNode && (startNode.nodeType === 3 ? startNode.parentElement : startNode)) || document.body;
    // Use our scoped content root instead of generic closest()
    const container = getContentRoot() || host.closest('article, [role="main"], main, body') || document.body;
    const { full: fullText } = linearizeTextNodes(container);
    
    if (!fullText) {
      console.warn('[ExtractContext] No text in container');
      return { before: '', after: '' };
    }

    const normalizeWS = (s) => normalizeTextForMatching(s || '');
    const normalizedHighlight = normalizeTextForMatching(highlightedText);
    const escaped = normalizedHighlight.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'i');

    const normalizedFull = normalizeTextForMatching(fullText);
    let match = normalizedFull.match(pattern);
    if (!match) {
      const idx = normalizedFull.indexOf(normalizedHighlight);
      if (idx >= 0) {
        match = { index: idx, 0: normalizedHighlight };
      }
    }
    if (!match) {
      const normFull = normalizeWS(fullText);
      const normHighlight = normalizeWS(normalizedHighlight);
      const idxNorm = normFull.indexOf(normHighlight);
      if (idxNorm !== -1) {
        // Approximate raw index by finding the first occurrence of the first token in raw text
        const firstToken = normHighlight.split(' ')[0] || '';
        const approx = fullText.toLowerCase().indexOf(firstToken.toLowerCase());
        if (approx !== -1) {
          match = { index: approx, 0: highlightedText };
        }
      }
    }

    if (!match || typeof match.index !== 'number') {
      console.warn('[ExtractContext] Text not found');
      return { before: '', after: '' };
    }
    
    const startIdx = match.index;
    const endIdx = startIdx + (match[0] ? match[0].length : highlightedText.length);
    const before = fullText.slice(Math.max(0, startIdx - 30), startIdx);
    const after = fullText.slice(endIdx, endIdx + 30);
    console.log(`  📍 Context: before(${before.length} chars) after(${after.length} chars)`);
    return { before, after };
  } catch (e) {
    console.warn('Context extraction error:', e);
    return { before: '', after: '' };
  }
}

// --- Helper: Get first tagId from URL (hash or search), supports multiples ---
function getTagIdFromUrl() {
  const hash = window.location.hash || '';
  const search = window.location.search || '';
  const combined = `${hash}&${search}`;
  const matches = [...combined.matchAll(/holorunTagId=([^&#]+)/g)];
  return matches.length ? decodeURIComponent(matches[0][1]) : null;
}

// --- Helper: Auto-jump to tagId after restore, with retry logic ---
function autoJumpIfTagIdPresent() {
  const tagId = getTagIdFromUrl();
  if (tagId) {
    console.log('[HolorunTagger] Auto-jump requested for tagId:', tagId);
    // Use jumpToHighlight which handles restoration if needed
    setTimeout(() => {
      jumpToHighlight(tagId);
    }, 1000);
  }
}


function removePopup() {
  const p = document.getElementById("tag-popup");
  if (p) {
    p.remove();
    console.log("Popup removed");
  }
}

/* ==================== SELECTION TRACKING ==================== */

// Debounced scroll handler: re-run restoreTags after user scrolls (helps lazy-loaded/older content)
function handleScrollForRestore() {
  if (scrollRestoreTimer) {
    clearTimeout(scrollRestoreTimer);
  }
  scrollRestoreTimer = setTimeout(() => {
    scrollRestoreTimer = null;
    console.log("🌐 Scroll detected, re-checking tags for newly loaded content...");
    restoreTags();
  }, 700);
}
window.addEventListener('scroll', handleScrollForRestore, { passive: true });

document.addEventListener("mouseup", (e) => {
  // Skip if clicking on our UI
  if (e.target && e.target.closest && e.target.closest('#tag-popup, #tag-panel')) {
    return;
  }
  const sel = window.getSelection();
  if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
    try {
      const range = sel.getRangeAt(0);
      
      // Check if selection contains or is inside non-text elements
      const isIconElement = (el) => {
        if (!el) return false;
        if (el.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"]')) return false;
        const hasLabel = (el.getAttribute && (el.getAttribute('aria-label') || el.getAttribute('title'))) || (el.dataset && el.dataset.icon);
        const looksLikeIcon = el.matches && el.matches('[class*="icon"], [data-icon], [role="img"], i, svg use');
        return Boolean(looksLikeIcon && hasLabel);
      };

      const isInvalidSelection = (node) => {
        if (!node) return false;
        const el = node.nodeType === 1 ? node : node.parentElement;
        if (!el) return false;
        // Allow icon selections when we can capture a label for restoration
        if (isIconElement(el)) return false;
        // Check if inside SVG, canvas, code blocks, or React managed content
        return el.closest('svg, canvas, pre, code, [class*="diagram"], .mermaid, [contenteditable], [role="textbox"], script, style');
      };
      
      // Validate both start and end containers
      if (isInvalidSelection(range.startContainer) || isInvalidSelection(range.endContainer)) {
        console.warn("⚠️ Selection is inside a non-text element (icon, diagram, code, or input). Tag regular text only.");
        return;
      }
      
      lastSelectionRange = range.cloneRange();
      console.log("✓ Selection saved:", sel.toString().substring(0, 60));
    } catch (err) {
      console.error("Selection error:", err);
    }
  }
});

/* ==================== APPLY PENDING HIGHLIGHT ==================== */

function applyPendingHighlight() {
  console.log("Applying highlight...");
  if (!lastSelectionRange) {
    console.error("No selection range available");
    alert("Please select text first!");
    return;
  }
  try {
    // Validate range contains only text nodes
    if (!lastSelectionRange.startContainer || !lastSelectionRange.startContainer.parentNode) {
      console.warn("Selection range is stale, cannot highlight");
      alert("Selection expired. Please select text again.");
      return;
    }
    
    const startEl = lastSelectionRange.startContainer.nodeType === 1 
      ? lastSelectionRange.startContainer 
      : lastSelectionRange.startContainer.parentElement;
    const endEl = lastSelectionRange.endContainer.nodeType === 1 
      ? lastSelectionRange.endContainer 
      : lastSelectionRange.endContainer.parentElement;
    
    // Reject if inside non-text elements
    const looksLikeIcon = (el) => {
      if (!el || !el.closest) return false;
      return el.closest('[class*="icon"], [data-icon], [role="img"], i') && !el.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"]');
    };

    if ((startEl && startEl.closest && startEl.closest('svg, canvas, pre, code, [class*="diagram"], .mermaid, [contenteditable], [role="textbox"]')) ||
        (endEl && endEl.closest && endEl.closest('svg, canvas, pre, code, [class*="diagram"], .mermaid, [contenteditable], [role="textbox"]'))) {
      // Allow icons with labels, otherwise block
      const startIconOk = startEl && looksLikeIcon(startEl);
      const endIconOk = endEl && looksLikeIcon(endEl);
      if (!startIconOk && !endIconOk) {
        alert("⚠️ Cannot tag text inside diagrams, code blocks, or input fields. Please select regular text only.");
        return;
      }
    }
    
    // Remove existing pending
    if (pendingHighlightSpan) {
      console.log("Removing old pending highlight");
      cancelTagging();
    }
    
    // IMPORTANT: Extract context BEFORE modifying DOM (before extractContents)
    const selectedText = lastSelectionRange.toString();
    const pendingContext = extractContext(
      lastSelectionRange.startContainer,
      lastSelectionRange.endContainer,
      selectedText,
      100
    );
    console.log(`  💾 Saved context for later: before="${pendingContext.before.substring(0, 30)}" after="${pendingContext.after.substring(0, 30)}"`);
    
    // Detect if selection is primarily an icon with a label we can store
    const detectIconTarget = () => {
      const el = startEl || endEl;
      if (!el || !el.closest) return null;
      const iconEl = el.closest('[class*="icon"], [data-icon], [role="img"], i');
      if (!iconEl || iconEl.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"]')) return null;
      const label = (iconEl.getAttribute('aria-label') || iconEl.getAttribute('title') || (iconEl.dataset && iconEl.dataset.icon) || '').trim();
      if (!label) return null;
      return { iconEl, label };
    };

    const iconTarget = detectIconTarget();
    if (iconTarget) {
      window.__pendingIconMeta = extractIconMeta(iconTarget.iconEl, iconTarget.label);
      console.log(`  🎯 Captured icon with label "${iconTarget.label}"`);
    } else {
      window.__pendingIconMeta = null;
    }

    const span = document.createElement("span");
    span.className = "pending-highlight";
    const fragment = lastSelectionRange.extractContents();
    span.appendChild(fragment);
    lastSelectionRange.insertNode(span);
    pendingHighlightSpan = span;
    
    // Store context globally so saveTag() can use it
    window.__pendingHighlightContext = pendingContext;
    
    window.getSelection().removeAllRanges();
    console.log("✓ Highlight applied:", span.textContent.substring(0, 60));
  } catch (e) {
    console.error("❌ Highlight failed:", e);
    alert("Failed to highlight. The page content may have changed. Try selecting again.");
  }
}

/* ==================== OPEN POPUP ==================== */

function openPopup() {
  console.log("Creating popup...");
  removePopup();

  const popup = document.createElement("div");
  popup.id = "tag-popup";
  popup.style.cssText = 'position: fixed !important; top: 20px !important; right: 20px !important; z-index: 2147483647 !important; background: #111 !important; color: #fff !important; padding: 16px !important; border-radius: 10px !important; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.6), 0 0 0 1px rgba(255, 255, 255, 0.1) !important; display: flex !important; flex-direction: column !important; gap: 10px !important; min-width: 220px !important; max-width: 320px !important; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;';
  popup.innerHTML = `
    <input id="tag-input" type="text" placeholder="Enter tag name" maxlength="50" style="width: 100% !important; padding: 10px 12px !important; border: 2px solid #333 !important; border-radius: 6px !important; background: #1a1a1a !important; color: #fff !important; font-size: 14px !important; outline: none !important; transition: all 0.2s ease !important; display: block !important; box-sizing: border-box !important;" />
    <div style="display: flex !important; gap: 8px !important; width: 100% !important;">
      <button id="save-tag" style="flex: 1 !important; padding: 10px 16px !important; border: none !important; border-radius: 6px !important; cursor: pointer !important; font-size: 14px !important; background: #1a73e8 !important; color: #fff !important; font-weight: 600 !important; transition: all 0.2s ease !important;">Save</button>
      <button id="cancel-tag" style="flex: 1 !important; padding: 10px 16px !important; border: none !important; border-radius: 6px !important; cursor: pointer !important; font-size: 14px !important; background: #5f6368 !important; color: #fff !important; font-weight: 600 !important; transition: all 0.2s ease !important;">Cancel</button>
    </div>
  `;
  document.body.appendChild(popup);
  console.log("✓ Popup added to DOM with inline styles");

  const input = document.getElementById("tag-input");
  const saveBtn = document.getElementById("save-tag");
  const cancelBtn = document.getElementById("cancel-tag");
  if (!input || !saveBtn || !cancelBtn) {
    console.error("Popup elements not found!");
    return;
  }
  input.focus();
  saveBtn.onclick = () => {
    console.log("Save button clicked");
    saveTag();
  };
  cancelBtn.onclick = () => {
    console.log("Cancel button clicked");
    cancelTagging();
  };
  input.onkeydown = (e) => {
    if (e.key === "Enter") {
      console.log("Enter pressed");
      saveTag();
    } else if (e.key === "Escape") {
      console.log("Escape pressed");
      cancelTagging();
    }
  };
  console.log("✓ Popup handlers attached");
}

/* ==================== MESSAGE LISTENER ==================== */

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  try {
    if (msg && msg.type === "OPEN_TAG_POPUP") {
      console.log("📩 OPEN_TAG_POPUP message received");
      applyPendingHighlight();
      openPopup();
      sendResponse && sendResponse({ success: true });
      return true;
    }
  } catch (err) {
    console.error("Message handling error:", err);
  }
  return false;
});

/* ==================== CANCEL ==================== */

function cancelTagging() {
  console.log("Cancelling...");
  
  try {
    if (pendingHighlightSpan && pendingHighlightSpan.parentNode) {
      const parent = pendingHighlightSpan.parentNode;
      while (pendingHighlightSpan.firstChild) {
        parent.insertBefore(pendingHighlightSpan.firstChild, pendingHighlightSpan);
      }
      parent.removeChild(pendingHighlightSpan);
      parent.normalize();
      console.log("✓ Pending highlight removed");
    }
  } catch (e) {
    console.error("Cancel error:", e);
  }

  pendingHighlightSpan = null;
  lastSelectionRange = null;
  removePopup();
}

/* ==================== SAVE TAG ==================== */

function saveTag() {
  console.log("Saving tag...");
  
  const input = document.getElementById("tag-input");
  if (!input) {
    console.error("Input not found");
    return;
  }

  const topic = input.value.trim();
  
  if (!topic) {
    console.log("Empty topic");
    input.focus();
    return;
  }
  
  if (!pendingHighlightSpan) {
    console.error("No pending highlight");
    alert("No text highlighted!");
    removePopup();
    return;
  }

  let text = pendingHighlightSpan.textContent;
  const tagId = crypto.randomUUID();

  // If selection was an icon, text may be empty; use captured label
  const iconMeta = window.__pendingIconMeta || null;
  if ((!text || !text.trim()) && iconMeta && iconMeta.label) {
    text = iconMeta.label;
  }

  // IMPORTANT: DO NOT normalize saved text - keep original exactly as user selected
  // Normalization happens during SEARCH (in findRangeByRawText), not during save
  // This preserves exact character positions so emoji/whitespace differences don't break position mapping
  console.log(`  ✅ Saving original text (no normalization): "${text.substring(0, 60)}"`);

  // VALIDATION: Check tag size limits BEFORE saving
  const MAX_TAG_LENGTH = 500; // Maximum chars for reliable restore (matches highlight validation)
  const MAX_SAFE_LENGTH = 1000; // Absolute max for token matching
  
  if (text.length > MAX_TAG_LENGTH) {
    const shouldContinue = confirm(
      `⚠️ WARNING: Selected text is ${text.length} chars (recommended max: ${MAX_TAG_LENGTH} chars).\n\n` +
      `Large tags may fail to restore, especially:\n` +
      `• Text in code blocks or diagrams\n` +
      `• Selections spanning many messages\n\n` +
      `Recommendation: Select a shorter, focused excerpt (50-${MAX_TAG_LENGTH} chars).\n\n` +
      `Continue anyway?`
    );
    if (!shouldContinue) {
      console.log("  🚫 User cancelled due to large selection");
      return;
    }
    
    if (text.length > MAX_SAFE_LENGTH) {
      alert(
        `❌ REJECTED: Selection is ${text.length} chars (max ${MAX_SAFE_LENGTH} chars).\n\n` +
        `This tag will NEVER restore.\n\n` +
        `Please select a much shorter excerpt.`
      );
      console.error(`  ❌ Tag rejected: ${text.length} chars exceeds ${MAX_SAFE_LENGTH} char limit`);
      return;
    }
  }
  
  // Don't truncate - keep original text for accurate restoration
  // Whitespace-tolerant matching handles multi-paragraph selections
  
  // Use context extracted during applyPendingHighlight (before DOM was modified)
  const context = window.__pendingHighlightContext || { before: '', after: '' };
  console.log(`  📍 Context: before="${context.before.substring(0, 40)}" after="${context.after.substring(0, 40)}"`);
  if (!context.before && !context.after) {
    console.warn(`  ⚠️ WARNING: No context captured! This tag may not restore correctly in changed conversations.`);
  }
  window.__pendingHighlightContext = null; // Clear after use
  window.__pendingIconMeta = null;

  // Clean URL: remove holorunTagId parameter before storing
  let cleanUrl = location.href;
  cleanUrl = cleanUrl.replace(/#holorunTagId=[^&]*&?/g, '#');
  cleanUrl = cleanUrl.replace(/\?holorunTagId=[^&]*&?/g, '?');
  cleanUrl = cleanUrl.replace(/#$/, '').replace(/\?$/, ''); // remove trailing # or ?

  const containerTag = getNearestBlockTag(pendingHighlightSpan);
  const kind = classifyTagKind(text, containerTag, iconMeta);

  const tag = {
    id: tagId,
    topic: topic,
    text: text,
    context: context,  // NEW: surrounding text for robust restoration
    url: cleanUrl,
    timestamp: Date.now(),
    iconMeta: iconMeta || null,
    kind: kind,
    containerTag: containerTag || ''
  };

  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("💾 SAVING TAG");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log(`📌 Tag ID: ${tagId}`);
  console.log(`🏷️  Topic: "${topic}"`);
  console.log(`📝 Text Type: ${kind}`);
  console.log(`📏 Text Length: ${text.length} chars`);
  console.log(`📍 Container: <${containerTag}>`);
  console.log(`🌐 URL: ${cleanUrl}`);
  console.log(`📄 Text Preview: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`); 
  console.log(`📍 Context Before: "${context.before.substring(0, 30)}"`);
  console.log(`📍 Context After: "${context.after.substring(0, 30)}"`);
  if (iconMeta) {
    console.log(`🎨 Icon Meta: label="${iconMeta.label}", ariaLabel="${iconMeta.ariaLabel}"`);
  }
  console.log("Full tag object:", tag);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");

  safeStorageGet({ tags: [] }, (res) => {
    const tags = res.tags || [];
    
    // Deduplication: Remove any older tags with the same text on the same page
    // This prevents duplicates and keeps newest highlight at top
    const newTags = tags.filter(t => !(samePageUrl(t.url, tag.url) && normalizeSimple(t.text) === normalizeSimple(text)));
    
    newTags.unshift(tag); // Add new tag to front
    
    safeStorageSet({ tags: newTags }, () => {
      if (chrome.runtime.lastError) {
        console.error("Storage error:", chrome.runtime.lastError);
        alert("Failed to save tag!");
      } else {
        console.log("✓ Saved! Total tags:", newTags.length);
        commitHighlight(topic, tagId);
        renderPanel();
        removePopup();
      }
    });
  });
}

/* ==================== JUMP TO HIGHLIGHT FUNCTION ==================== */

async function jumpToHighlight(tagId) {
  console.log("🎯 JUMPING TO:", tagId);
  lastJumpedTagId = tagId; // Track this jump to avoid re-searching from top on next jump
  
  // Ensure content is visible before attempting to find and style the highlight
  await expandCollapsedChatGPTMessages();
  
  // Prefer React-safe temporary styling with clear visual pulse
  // Use globally exposed function to avoid scope issues
  try {
    if (window.__applyHighlightStyling) {
      await window.__applyHighlightStyling(tagId);
    } else {
      throw new Error('applyHighlightStyling is not available');
    }
  } catch (e) {
    console.warn('[Jump] Styling failed unexpectedly, attempting restore/pulse fallback:', e.message);
    // Fallback only if styling path throws (should be rare)
    restoreTagHighlight(tagId, { pulse: true, scroll: true });
  }
}

/* ==================== DEEP LINKING SUPPORT ==================== */

function handleDeepLink() {
  const tagId = getHashParam('holorunTagId');
  if (!tagId) return;
  
  console.log(`[DeepLink] Attempting to jump to tag from URL hash: ${tagId}`);
  
  // Retry with backoff in case highlight isn't ready yet
  const startTime = Date.now();
  const maxRetries = 5;
  let retryCount = 0;
  
  const tryJump = () => {
    const highlight = document.querySelector(`[data-tag-id="${tagId}"]`);
    if (highlight) {
      console.log(`[DeepLink] ✅ Found highlight, jumping to it`);
      jumpToHighlight(tagId);
      // Clean URL after successful jump
      setTimeout(() => {
        const cleanedUrl = location.href
          .replace(/#holorunTagId=[^&]*/g, '')
          .replace(/\?holorunTagId=[^&]*/g, '')
          .replace(/#$/, '')
          .replace(/\?$/, '');
        history.replaceState(null, '', cleanedUrl);
      }, 100);
      return;
    }
    
    retryCount++;
    if (retryCount < maxRetries && Date.now() - startTime < 5000) {
      console.log(`[DeepLink] Highlight not found yet, retrying (${retryCount}/${maxRetries})...`);
      setTimeout(tryJump, 200 * retryCount); // Exponential backoff
    } else {
      console.warn(`[DeepLink] ❌ Could not find highlight after ${retryCount} retries`);
    }
  };
  
  tryJump();
}

/* ==================== COMMIT HIGHLIGHT ==================== */

function truncateToSingleParagraph(text) {
  // Split on double newlines (paragraphs) or single newlines followed by significant content
  const lines = text.split(/\n+/);
  const firstParagraph = lines[0].trim();
  
  // If first paragraph is very short (< 20 chars), try to include more context
  if (firstParagraph.length < 20 && lines.length > 1) {
    return (firstParagraph + ' ' + (lines[1] || '')).trim().substring(0, 200);
  }
  
  // Otherwise use first paragraph, max 200 chars for safety
  return firstParagraph.substring(0, 200);
}

// Choose a safe snippet for matching to avoid giant selections slowing or breaking restore
function getMatchableText(tag) {
  const text = (tag && tag.text) ? tag.text : '';
  const MAX_MATCHABLE_LENGTH = 600; // keep in sync with highlight safety limits
  if (text.length <= MAX_MATCHABLE_LENGTH) return text;

  const snippet = text.substring(0, MAX_MATCHABLE_LENGTH);
  console.warn(`[MatchText] ⚠️ Tag text is ${text.length} chars; matching with first ${snippet.length} chars. Re-save with a smaller selection for best accuracy.`);
  return snippet;
}

/**
 * One-time migration: Truncate all legacy multi-paragraph tags
 * This runs once to fix tags saved before truncation was implemented
 */
function migrateLegacyTags() {
  return new Promise((resolve) => {
    const MIGRATION_FLAG = 'holorunTagsMigrated';
    // If extension context is invalid, skip migration and mark flag in localStorage to avoid repeats
    if (!isExtensionContextValid()) {
      try { localStorage.setItem(MIGRATION_FLAG, 'true'); } catch (_) {}
      console.log('↷ Skipping legacy tag migration (context invalid)');
      resolve();
      return;
    }
    
    // Check if already migrated
    safeStorageGet({ [MIGRATION_FLAG]: false }, (migrated) => {
      if (migrated[MIGRATION_FLAG]) {
        console.log('✓ Legacy tags already migrated');
        resolve();
        return;
      }

      console.log('🔄 Migrating legacy multi-paragraph tags...');
      
      safeStorageGet({ tags: [] }, (result) => {
        const tags = result.tags || [];
        let migrationCount = 0;

        const migratedTags = tags.map(tag => {
          const originalLength = tag.text.length;
          const truncated = truncateToSingleParagraph(tag.text);
          
          if (truncated !== tag.text) {
            migrationCount++;
            console.log(`  📝 Migrated "${tag.topic}": ${originalLength} → ${truncated.length} chars`);
            return { ...tag, text: truncated };
          }
          return tag;
        });

        if (migrationCount > 0) {
          safeStorageSet({ tags: migratedTags }, () => {
            console.log(`✅ Migrated ${migrationCount} tag(s)`);
            safeStorageSet({ [MIGRATION_FLAG]: true }, () => resolve());
          });
        } else {
          console.log('✓ No tags needed migration');
          safeStorageSet({ [MIGRATION_FLAG]: true }, () => resolve());
        }
      });
    });
  });
}

function commitHighlight(tagText, tagId) {
  console.log("Committing highlight:", tagId);
  
  if (!pendingHighlightSpan) {
    console.error("No pending highlight to commit");
    return;
  }

  pendingHighlightSpan.className = "tag-highlight";
  pendingHighlightSpan.dataset.tagId = tagId;
  pendingHighlightSpan.dataset.tag = tagText;
  
  // Log CSS styles applied
  const styles = window.getComputedStyle(pendingHighlightSpan);
  console.log(`[Highlight CSS] background: ${styles.backgroundColor}`);
  console.log(`[Highlight CSS] padding: ${styles.padding}`);
  console.log(`[Highlight CSS] border-radius: ${styles.borderRadius}`);
  console.log(`[Highlight CSS] display: ${styles.display}`);
  console.log(`[Highlight CSS] Text content: "${pendingHighlightSpan.textContent.substring(0, 40)}..."`);
  
  // Add click handler - use the jumpToHighlight function
  pendingHighlightSpan.onclick = function(e) {
    e.preventDefault();
    e.stopPropagation();
    console.log("🖱 Highlight clicked! Tag ID:", tagId);
    jumpToHighlight(tagId);
  };

  console.log("✓ Highlight committed");
  pendingHighlightSpan = null;
  lastSelectionRange = null;
}

/* ==================== REMOVE HIGHLIGHT ==================== */

function removeHighlight(tagId) {
  console.log("Removing highlight:", tagId);
  
  const span = document.querySelector(`[data-tag-id="${tagId}"]`);
  if (!span) {
    console.warn("Highlight not found");
    return;
  }

  const parent = span.parentNode;
  while (span.firstChild) {
    parent.insertBefore(span.firstChild, span);
  }
  parent.removeChild(span);
  parent.normalize();
  
  console.log("✓ Highlight removed");
}

/* ==================== RENDER PANEL ==================== */

// Helper to get count of open Chrome tabs
function getOpenTabCount(callback) {
  if (!chrome.runtime) {
    callback(0);
    return;
  }
  
  try {
    chrome.runtime.sendMessage({ type: 'GET_TAB_COUNT' }, (response) => {
      if (chrome.runtime.lastError) {
        console.warn('[TabCount] Error getting tab count:', chrome.runtime.lastError);
        callback(0);
      } else {
        callback(response?.count || 0);
      }
    });
  } catch (e) {
    console.warn('[TabCount] Exception getting tab count:', e);
    callback(0);
  }
}

const ASSISTANT_STORAGE_KEY = 'holorunAssistantEnabled';
let assistantEnabledCache = true;
let lastAssistantActivitySent = 0;

function loadAssistantEnabled(callback) {
  try {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      chrome.storage.local.get([ASSISTANT_STORAGE_KEY], (res) => {
        const value = res && typeof res[ASSISTANT_STORAGE_KEY] === 'boolean'
          ? res[ASSISTANT_STORAGE_KEY]
          : true;
        assistantEnabledCache = value;
        if (typeof callback === 'function') callback(value);
      });
      return;
    }
  } catch {
    // Ignore and fallback.
  }

  try {
    const stored = localStorage.getItem(ASSISTANT_STORAGE_KEY);
    const value = stored === null ? true : stored === 'true';
    assistantEnabledCache = value;
    if (typeof callback === 'function') callback(value);
  } catch {
    assistantEnabledCache = true;
    if (typeof callback === 'function') callback(true);
  }
}

function setAssistantEnabled(value, callback) {
  assistantEnabledCache = Boolean(value);
  try {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      chrome.storage.local.set({ [ASSISTANT_STORAGE_KEY]: assistantEnabledCache }, () => {
        if (typeof callback === 'function') callback(assistantEnabledCache);
      });
      return;
    }
  } catch {
    // Ignore and fallback.
  }

  try {
    localStorage.setItem(ASSISTANT_STORAGE_KEY, assistantEnabledCache ? 'true' : 'false');
  } catch {
    // Ignore storage errors.
  }

  if (typeof callback === 'function') callback(assistantEnabledCache);
}

if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.onChanged) {
  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== 'local') return;
    if (changes[ASSISTANT_STORAGE_KEY]) {
      assistantEnabledCache = Boolean(changes[ASSISTANT_STORAGE_KEY].newValue);
    }
  });
}

loadAssistantEnabled();

function isHoloUiElement(node) {
  try {
    const el = node && node.nodeType === 1 ? node : node && node.parentElement ? node.parentElement : null;
    if (!el || typeof el.closest !== 'function') return false;
    return Boolean(
      el.closest('#tag-panel-host') ||
      el.closest('#tag-fab-host') ||
      el.closest('#tag-panel') ||
      el.closest('#tag-popup') ||
      el.closest('.holorun-popup')
    );
  } catch {
    return false;
  }
}

document.addEventListener('click', (event) => {
  if (!assistantEnabledCache) return;
  if (isHoloUiElement(event.target)) return;
  const now = Date.now();
  if (now - lastAssistantActivitySent < 1000) return;
  lastAssistantActivitySent = now;

  try {
    if (typeof holoDiagLog === 'function') {
      holoDiagLog('assistant.activity', { action: 'click', timestamp: now });
    }
    if (chrome.runtime) {
      chrome.runtime.sendMessage({ type: 'TAB_ACTIVITY', action: 'click', timestamp: now });
    }
  } catch {
    // Ignore send failures.
  }
}, true);

let tabCountIntervalId = null;

function updatePanelTabCount() {
  const panelHost = document.getElementById('tag-panel-host');
  if (!panelHost || !panelHost.shadowRoot) return;

  const countEl = panelHost.shadowRoot.querySelector('#tab-count');
  if (!countEl) return;

  getOpenTabCount((count) => {
    countEl.textContent = `Tabs: ${count}`;
  });
}

/* ==================== ORGANIZER SELECTION TRACKING ==================== */

let lastPointerElement = null;
document.addEventListener('pointerdown', (event) => {
  lastPointerElement = event.target || null;
}, true);

/* ==================== TAB INTERFACE CONTROL SYSTEM ==================== */

// Tab management state
let currentActiveTab = 'all-tags';
let tabsData = {
  'all-tags': { name: 'All Tags', icon: '📋', content: null },
  'recent': { name: 'Recent', icon: '🕒', content: null },
  'favorites': { name: 'Favorites', icon: '⭐', content: null },
  'organizer': { name: 'Organizer', icon: '🧩', content: null },
  'workflow': { name: 'Workflow', icon: '🧭', content: null },
  'assistant': { name: 'Assistant', icon: '🤖', content: null },
  'settings': { name: 'Settings', icon: '⚙️', content: null }
};

// Tab interface control functions
const TabInterface = {
  
  // Create a new tab
  createTab: function(tabId, tabData) {
    if (!tabId || !tabData) {
      console.error('[TabInterface] createTab: Invalid parameters');
      return false;
    }
    
    tabsData[tabId] = {
      name: tabData.name || 'Untitled',
      icon: tabData.icon || '📄',
      content: tabData.content || null,
      visible: tabData.visible !== false,
      disabled: tabData.disabled || false
    };
    
    console.log(`[TabInterface] Created tab: ${tabId}`);
    this.refreshTabs();
    return true;
  },
  
  // Remove a tab
  removeTab: function(tabId) {
    if (!tabsData[tabId]) {
      console.error(`[TabInterface] removeTab: Tab ${tabId} does not exist`);
      return false;
    }
    
    // Don't allow removing the default tab
    if (tabId === 'all-tags') {
      console.warn('[TabInterface] Cannot remove default tab');
      return false;
    }
    
    delete tabsData[tabId];
    
    // Switch to default tab if current tab was removed
    if (currentActiveTab === tabId) {
      this.switchTab('all-tags');
    }
    
    console.log(`[TabInterface] Removed tab: ${tabId}`);
    this.refreshTabs();
    return true;
  },
  
  // Switch to a specific tab
  switchTab: function(tabId) {
    if (!tabsData[tabId]) {
      console.error(`[TabInterface] switchTab: Tab ${tabId} does not exist`);
      return false;
    }
    
    if (tabsData[tabId].disabled) {
      console.warn(`[TabInterface] Tab ${tabId} is disabled`);
      return false;
    }
    
    const previousTab = currentActiveTab;
    currentActiveTab = tabId;
    
    // Update tab visual states
    this.updateTabStates();
    
    // Load tab content
    this.loadTabContent(tabId);
    
    console.log(`[TabInterface] Switched from ${previousTab} to ${tabId}`);
    return true;
  },
  
  // Update visual states of tabs
  updateTabStates: function() {
    const panelHost = document.getElementById("tag-panel-host");
    if (!panelHost || !panelHost.shadowRoot) return;
    
    const tabButtons = panelHost.shadowRoot.querySelectorAll('.tab-button');
    tabButtons.forEach(button => {
      const tabId = button.dataset.tabId;
      if (tabId === currentActiveTab) {
        button.classList.add('active');
      } else {
        button.classList.remove('active');
      }
    });
  },
  
  // Load content for specific tab
  loadTabContent: function(tabId) {
    const panelHost = document.getElementById("tag-panel-host");
    if (!panelHost || !panelHost.shadowRoot) return;
    
    const contentArea = panelHost.shadowRoot.querySelector('#tab-content');
    if (!contentArea) return;
    
    // Clear current content
    contentArea.innerHTML = '';
    
    switch(tabId) {
      case 'all-tags':
        this.renderAllTagsContent(contentArea);
        break;
      case 'recent':
        this.renderRecentContent(contentArea);
        break;
      case 'favorites':
        this.renderFavoritesContent(contentArea);
        break;
      case 'organizer':
        this.renderOrganizerContent(contentArea);
        break;
      case 'workflow':
        this.renderWorkflowContent(contentArea);
        break;
      case 'assistant':
        this.renderAssistantContent(contentArea);
        break;
      case 'settings':
        this.renderSettingsContent(contentArea);
        break;
      default:
        // Custom tab content
        if (tabsData[tabId].content) {
          if (typeof tabsData[tabId].content === 'function') {
            tabsData[tabId].content(contentArea);
          } else {
            contentArea.innerHTML = tabsData[tabId].content;
          }
        }
    }
  },
  
  // Render all tags content (default)
  renderAllTagsContent: function(container) {
    safeStorageGet({ tags: [] }, (res) => {
      const tags = res.tags || [];
      tags.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
      
      if (tags.length === 0) {
        container.innerHTML = '<div class="empty-state">No tags yet. Highlight some text and create your first tag!</div>';
        return;
      }
      
      const tagsList = document.createElement('div');
      tagsList.className = 'tags-list';
      
      tags.forEach(tag => {
        const tagRow = this.createTagElement(tag);
        tagsList.appendChild(tagRow);
      });
      
      container.appendChild(tagsList);
    });
  },
  
  // Render recent content
  renderRecentContent: function(container) {
    safeStorageGet({ tags: [] }, (res) => {
      const tags = res.tags || [];
      const recentTags = tags
        .filter(tag => {
          const daysSince = (Date.now() - (tag.timestamp || 0)) / (1000 * 60 * 60 * 24);
          return daysSince <= 7; // Last 7 days
        })
        .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))
        .slice(0, 20); // Top 20 recent
        
      if (recentTags.length === 0) {
        container.innerHTML = '<div class="empty-state">No recent tags from the last 7 days.</div>';
        return;
      }
      
      const tagsList = document.createElement('div');
      tagsList.className = 'tags-list';
      
      recentTags.forEach(tag => {
        const tagRow = this.createTagElement(tag);
        tagsList.appendChild(tagRow);
      });
      
      container.appendChild(tagsList);
    });
  },
  
  // Render favorites content
  renderFavoritesContent: function(container) {
    safeStorageGet({ tags: [], favorites: [] }, (res) => {
      const tags = res.tags || [];
      const favorites = res.favorites || [];
      const favoriteTags = tags.filter(tag => favorites.includes(tag.id));
      
      if (favoriteTags.length === 0) {
        container.innerHTML = '<div class="empty-state">No favorite tags yet. Star some tags to see them here!</div>';
        return;
      }
      
      const tagsList = document.createElement('div');
      tagsList.className = 'tags-list';
      
      favoriteTags.forEach(tag => {
        const tagRow = this.createTagElement(tag, true);
        tagsList.appendChild(tagRow);
      });
      
      container.appendChild(tagsList);
    });
  },
  
  // Render settings content
  renderSettingsContent: function(container) {
    container.innerHTML = `
      <div class="settings-content">
        <div class="setting-group">
          <h3>Display Options</h3>
          <label class="setting-item">
            <input type="checkbox" id="show-timestamps"> Show timestamps
          </label>
          <label class="setting-item">
            <input type="checkbox" id="compact-view"> Compact view
          </label>
        </div>
        <div class="setting-group">
          <h3>Data Management</h3>
          <button id="export-tags" class="setting-button">Export All Tags</button>
          <button id="clear-all" class="setting-button danger">Clear All Tags</button>
        </div>
        <div class="setting-group">
          <h3>About</h3>
          <p>HoloTagg Extension v1.0.0</p>
          <p>Highlight, tag, and revisit text on any webpage.</p>
        </div>
      </div>
    `;
    
    // Add event listeners for settings
    const showTimestamps = container.querySelector('#show-timestamps');
    const compactView = container.querySelector('#compact-view');
    const exportBtn = container.querySelector('#export-tags');
    const clearBtn = container.querySelector('#clear-all');
    
    if (exportBtn) {
      exportBtn.onclick = () => this.exportTags();
    }
    
    if (clearBtn) {
      clearBtn.onclick = () => this.clearAllTags();
    }
  },

  // Render organizer content
  renderOrganizerContent: function(container) {
    OrganizerManager.renderOrganizerContent(container);
  },

  // Render workflow content
  renderWorkflowContent: function(container) {
    container.innerHTML = `
      <div class="settings-content">
        <div class="setting-group">
          <h3>Workflow Snapshot</h3>
          <div class="workflow-stats" id="workflow-stats">Loading tab metrics…</div>
          <div class="workflow-note">Grid view uses tab metadata and favicons (no live thumbnails).</div>
        </div>
        <div class="setting-group">
          <h3>Tab Controls</h3>
          <div class="workflow-controls">
            <input id="workflow-search" class="organizer-input" placeholder="Search tabs by title or URL" />
            <button id="workflow-refresh" class="setting-button">Refresh</button>
          </div>
          <div class="workflow-actions" id="workflow-actions">
            <input id="workflow-group-name" class="organizer-input" placeholder="Group name (optional)" />
            <select id="workflow-group-color" class="organizer-select small">
              <option value="" selected>Group color</option>
              <option value="grey">Grey</option>
              <option value="blue">Blue</option>
              <option value="red">Red</option>
              <option value="yellow">Yellow</option>
              <option value="green">Green</option>
              <option value="pink">Pink</option>
              <option value="purple">Purple</option>
              <option value="cyan">Cyan</option>
              <option value="orange">Orange</option>
            </select>
            <button id="workflow-group" class="setting-button primary">Group Selected</button>
            <button id="workflow-ungroup" class="setting-button">Ungroup Selected</button>
            <button id="workflow-move" class="setting-button">Move Selected to Front</button>
            <button id="workflow-close" class="setting-button danger">Close Selected</button>
          </div>
        </div>
        <div class="setting-group">
          <h3>All Tabs</h3>
          <div class="workflow-grid" id="workflow-grid"></div>
        </div>
      </div>
    `;

    if (!window.chromeTabController) {
      console.error('[Workflow] ChromeTabController not available');
      const stats = container.querySelector('#workflow-stats');
      if (stats) stats.textContent = 'Tab controls unavailable in this context.';
      return;
    }
    
    console.log('[Workflow] ChromeTabController available, initializing UI');

    const statsEl = container.querySelector('#workflow-stats');
    const gridEl = container.querySelector('#workflow-grid');
    const searchInput = container.querySelector('#workflow-search');
    const refreshBtn = container.querySelector('#workflow-refresh');
    const groupBtn = container.querySelector('#workflow-group');
    const ungroupBtn = container.querySelector('#workflow-ungroup');
    const moveBtn = container.querySelector('#workflow-move');
    const closeBtn = container.querySelector('#workflow-close');
    const groupNameInput = container.querySelector('#workflow-group-name');
    const groupColorSelect = container.querySelector('#workflow-group-color');

    console.log('[Workflow] UI elements:', {
      statsEl: !!statsEl,
      gridEl: !!gridEl,
      searchInput: !!searchInput,
      refreshBtn: !!refreshBtn
    });

    let currentTabs = [];

    const renderStats = (tabs) => {
      if (!statsEl) return;
      const total = tabs.length;
      const pinned = tabs.filter(tab => tab.pinned).length;
      const audible = tabs.filter(tab => tab.audible).length;
      const grouped = tabs.filter(tab => tab.groupId && tab.groupId !== -1).length;
      const active = tabs.find(tab => tab.active);

      const domainCounts = {};
      tabs.forEach(tab => {
        try {
          const host = new URL(tab.url).hostname;
          domainCounts[host] = (domainCounts[host] || 0) + 1;
        } catch {
          // ignore
        }
      });
      const topDomains = Object.entries(domainCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 4)
        .map(([domain, count]) => `${domain} (${count})`)
        .join(' • ');

      statsEl.textContent = `Total: ${total} • Pinned: ${pinned} • Grouped: ${grouped} • Audible: ${audible}`;
      if (active) {
        const activeSpan = document.createElement('div');
        activeSpan.className = 'workflow-meta';
        activeSpan.textContent = `Active: ${active.title || active.url}`;
        statsEl.appendChild(activeSpan);
      }
      if (topDomains) {
        const domainSpan = document.createElement('div');
        domainSpan.className = 'workflow-meta';
        domainSpan.textContent = `Top domains: ${topDomains}`;
        statsEl.appendChild(domainSpan);
      }
    };

    const getHueFromString = (value) => {
      const str = String(value || 'tab');
      let hash = 0;
      for (let i = 0; i < str.length; i += 1) {
        hash = (hash << 5) - hash + str.charCodeAt(i);
        hash |= 0;
      }
      return Math.abs(hash) % 360;
    };

    const createGridCard = (tab) => {
      const card = document.createElement('button');
      card.className = `workflow-card${tab.active ? ' active' : ''}`;
      card.type = 'button';
      card.title = tab.url || tab.title || 'Untitled tab';
      card.dataset.tabId = tab.id;

      const chrome = document.createElement('div');
      chrome.className = 'workflow-card-chrome';

      const favicon = document.createElement('img');
      favicon.className = 'workflow-card-favicon';
      favicon.src = tab.favIconUrl || `https://www.google.com/s2/favicons?domain=${(() => {
        try {
          return new URL(tab.url).hostname;
        } catch {
          return '';
        }
      })()}`;
      favicon.alt = '';
      favicon.onerror = () => { favicon.style.display = 'none'; };

      const domain = document.createElement('div');
      domain.className = 'workflow-card-domain';
      domain.textContent = (() => {
        try {
          return new URL(tab.url).hostname;
        } catch {
          return 'unknown';
        }
      })();

      const closeBtn = document.createElement('button');
      closeBtn.className = 'workflow-card-close';
      closeBtn.type = 'button';
      closeBtn.textContent = '×';
      closeBtn.title = 'Close tab';
      closeBtn.onclick = async (event) => {
        event.stopPropagation();
        await ChromeTabController.closeTab(tab.id);
        refreshTabs();
      };

      chrome.appendChild(favicon);
      chrome.appendChild(domain);
      chrome.appendChild(closeBtn);

      const preview = document.createElement('div');
      preview.className = 'workflow-card-preview';
      const hue = getHueFromString(tab.url || tab.title || tab.id);
      preview.style.background = `linear-gradient(135deg, hsl(${hue}, 62%, 78%) 0%, hsl(${(hue + 40) % 360}, 60%, 62%) 100%)`;

      const title = document.createElement('div');
      title.className = 'workflow-card-title';
      title.textContent = tab.title || tab.url || 'Untitled tab';

      const meta = document.createElement('div');
      meta.className = 'workflow-card-meta';
      
      // Add position indicator
      const positionBadge = document.createElement('span');
      positionBadge.className = 'workflow-card-position';
      positionBadge.textContent = `#${tab.index + 1}`;
      positionBadge.title = `Position ${tab.index + 1} in window ${tab.windowId}`;
      positionBadge.style.cssText = 'font-weight: 600; margin-right: 8px; color: var(--color-primary);';
      meta.appendChild(positionBadge);
      
      const flags = [tab.pinned ? 'Pinned' : null, tab.audible ? 'Audio' : null, tab.mutedInfo?.muted ? 'Muted' : null]
        .filter(Boolean)
        .join(' • ');
      
      const flagText = document.createTextNode(flags || 'Ready');
      meta.appendChild(flagText);

      card.appendChild(chrome);
      card.appendChild(preview);
      card.appendChild(title);
      card.appendChild(meta);

      card.onclick = () => ChromeTabController.switchToTab(tab.id);
      return card;
    };

    const renderGrid = (tabs, query = '') => {
      console.log('[Workflow Grid] renderGrid called with', tabs.length, 'tabs');
      if (!gridEl) {
        console.error('[Workflow Grid] gridEl not found');
        return;
      }
      gridEl.innerHTML = '';
      const lowered = query.trim().toLowerCase();
      const filtered = lowered
        ? tabs.filter(tab => (tab.title || '').toLowerCase().includes(lowered) || (tab.url || '').toLowerCase().includes(lowered))
        : tabs;

      console.log('[Workflow Grid] Filtered to', filtered.length, 'tabs');

      if (!filtered.length) {
        const empty = document.createElement('div');
        empty.className = 'empty-state';
        empty.textContent = lowered ? 'No tabs match that search.' : 'No tabs available.';
        gridEl.appendChild(empty);
        return;
      }

      // Sort tabs by window ID, then by index (their actual position)
      const sorted = filtered.slice().sort((a, b) => {
        if (a.windowId !== b.windowId) {
          return a.windowId - b.windowId;
        }
        return a.index - b.index;
      });

      console.log('[Workflow Grid] Creating', sorted.length, 'cards sorted by window & position');

      // Group by window
      let currentWindowId = null;
      sorted.forEach(tab => {
        // Add window separator if we're entering a new window
        if (tab.windowId !== currentWindowId) {
          currentWindowId = tab.windowId;
          const windowLabel = document.createElement('div');
          windowLabel.className = 'workflow-window-label';
          windowLabel.textContent = `Window ${tab.windowId}`;
          gridEl.appendChild(windowLabel);
        }
        
        const card = createGridCard(tab);
        gridEl.appendChild(card);
      });

      console.log('[Workflow Grid] Grid rendering complete');
    };



    const refreshTabs = async () => {
      console.log('[Workflow] refreshTabs called');
      if (!statsEl || !gridEl) {
        console.error('[Workflow] statsEl or gridEl not found');
        return;
      }
      statsEl.textContent = 'Refreshing…';
      try {
        currentTabs = await ChromeTabController.getAllTabs();
        console.log('[Workflow] Got', currentTabs.length, 'tabs from ChromeTabController');
        renderStats(currentTabs);
        const query = searchInput ? searchInput.value : '';
        renderGrid(currentTabs, query);
      } catch (err) {
        console.error('[Workflow] refreshTabs error:', err);
        statsEl.textContent = 'Error loading tabs: ' + err.message;
      }
    };

    if (refreshBtn) {
      refreshBtn.onclick = () => refreshTabs();
    }

    if (searchInput) {
      searchInput.oninput = () => {
        renderGrid(currentTabs, searchInput.value);
      };
    }

    if (groupBtn) {
      groupBtn.onclick = async () => {
        const selected = getSelectedTabIds();
        if (!selected.length) {
          alert('Select at least one tab to group.');
          return;
        }
        const title = groupNameInput ? groupNameInput.value.trim() : '';
        const color = groupColorSelect ? groupColorSelect.value : '';
        await ChromeTabController.groupTabs(selected, {
          ...(title ? { title } : {}),
          ...(color ? { color } : {})
        });
        refreshTabs();
      };
    }

    if (ungroupBtn) {
      ungroupBtn.onclick = async () => {
        const selected = getSelectedTabIds();
        if (!selected.length) {
          alert('Select at least one tab to ungroup.');
          return;
        }
        await ChromeTabController.ungroupTabs(selected);
        refreshTabs();
      };
    }

    const getSelectedTabIds = () => {
      if (!gridEl) return [];
      const cards = gridEl.querySelectorAll('.workflow-card.selected');
      return Array.from(cards).map(card => Number(card.dataset.tabId)).filter(id => Number.isFinite(id));
    };

    if (moveBtn) {
      moveBtn.onclick = async () => {
        const selected = getSelectedTabIds();
        if (!selected.length) {
          alert('Select tabs to move.');
          return;
        }
        let index = 0;
        for (const tabId of selected) {
          await ChromeTabController.moveTabToPosition(tabId, index);
          index += 1;
        }
        refreshTabs();
      };
    }

    if (closeBtn) {
      closeBtn.onclick = async () => {
        const selected = getSelectedTabIds();
        if (!selected.length) {
          alert('Select tabs to close.');
          return;
        }
        if (!confirm(`Close ${selected.length} tab(s)?`)) return;
        await ChromeTabController.closeTab(selected);
        refreshTabs();
      };
    }

    refreshTabs();
  },

  renderAssistantContent: function(container) {
    container.innerHTML = `
      <div class="settings-content">
        <div class="setting-group">
          <h3>Assistant</h3>
          <div class="assistant-controls">
            <button id="assistant-toggle" class="setting-button">Assistant: --</button>
            <button id="assistant-refresh" class="setting-button">Refresh</button>
            <button id="assistant-center" class="setting-button">Center Panel</button>
          </div>
          <div class="assistant-note">Manual refresh only. No auto switching.</div>
        </div>
        <div class="setting-group">
          <h3>Relevant Tabs</h3>
          <div class="assistant-list" id="assistant-tabs"></div>
        </div>
        <div class="setting-group">
          <h3>Related Tags</h3>
          <div class="assistant-list" id="assistant-tags"></div>
        </div>
      </div>
    `;

    const toggleBtn = container.querySelector('#assistant-toggle');
    const refreshBtn = container.querySelector('#assistant-refresh');
    const centerBtn = container.querySelector('#assistant-center');
    const tabsList = container.querySelector('#assistant-tabs');
    const tagsList = container.querySelector('#assistant-tags');

    const updateToggle = (enabled) => {
      if (!toggleBtn) return;
      toggleBtn.textContent = `Assistant: ${enabled ? 'On' : 'Off'}`;
      if (enabled) {
        toggleBtn.classList.add('primary');
      } else {
        toggleBtn.classList.remove('primary');
      }
    };

    const centerPanel = () => {
      const panelHost = document.getElementById('tag-panel-host');
      if (!panelHost) return;
        panelHost.style.cssText = 'position: fixed !important; left: 50% !important; top: 50% !important; transform: translate(-50%, -50%) !important; z-index: 2147483647 !important; display: block !important; width: 600px !important; max-width: calc(100vw - 32px) !important; max-height: 80vh !important; cursor: grab !important;';

        // Add draggable logic
        let isDragging = false;
        let dragOffsetX = 0;
        let dragOffsetY = 0;
        let lastPanelPosition = { left: null, top: null };

        panelHost.onmousedown = function(e) {
          if (e.target !== panelHost) return;
          isDragging = true;
          panelHost.style.cursor = 'grabbing';
          const rect = panelHost.getBoundingClientRect();
          dragOffsetX = e.clientX - rect.left;
          dragOffsetY = e.clientY - rect.top;
        };
        document.onmousemove = function(e) {
          if (!isDragging) return;
          let left = e.clientX - dragOffsetX;
          let top = e.clientY - dragOffsetY;
          // Clamp within viewport
          left = Math.max(0, Math.min(left, window.innerWidth - panelHost.offsetWidth));
          top = Math.max(0, Math.min(top, window.innerHeight - panelHost.offsetHeight));
          panelHost.style.left = left + 'px';
          panelHost.style.top = top + 'px';
          panelHost.style.transform = 'none';
          lastPanelPosition.left = left;
          lastPanelPosition.top = top;
        };
        document.onmouseup = function() {
          if (isDragging) {
            isDragging = false;
            panelHost.style.cursor = 'grab';
          }
        };
        // Restore last position if available
        if (lastPanelPosition.left !== null && lastPanelPosition.top !== null) {
          panelHost.style.left = lastPanelPosition.left + 'px';
          panelHost.style.top = lastPanelPosition.top + 'px';
          panelHost.style.transform = 'none';
        }
      if (typeof holoDiagLog === 'function') {
        holoDiagLog('assistant.center_panel');
      }
    };

    const renderTabs = (tabs, activityMap = {}, activeTabId = null) => {
      if (!tabsList) return;
      tabsList.innerHTML = '';

      if (!tabs.length) {
        const empty = document.createElement('div');
        empty.className = 'empty-state';
        empty.textContent = 'No tabs available.';
        tabsList.appendChild(empty);
        return;
      }

      tabs.forEach(tab => {
        const row = document.createElement('div');
        row.className = 'assistant-row';

        const title = document.createElement('div');
        title.className = 'assistant-title';
        title.textContent = tab.title || tab.url || 'Untitled tab';

        const meta = document.createElement('div');
        meta.className = 'assistant-meta';
        const host = (() => {
          try {
            return new URL(tab.url).hostname;
          } catch {
            return 'unknown';
          }
        })();

        const activity = activityMap[String(tab.id)] || {};
        const clickCount = Number(activity.clickCount || 0);
        const activeLabel = tab.id === activeTabId ? 'Active' : null;
        const parts = [host, activeLabel, clickCount ? `Clicks: ${clickCount}` : null].filter(Boolean);
        meta.textContent = parts.join(' • ');

        const actions = document.createElement('div');
        actions.className = 'assistant-actions';
        const switchBtn = document.createElement('button');
        switchBtn.className = 'setting-button small';
        switchBtn.textContent = 'Switch';
        switchBtn.onclick = () => ChromeTabController.switchToTab(tab.id);
        actions.appendChild(switchBtn);

        row.appendChild(title);
        row.appendChild(meta);
        row.appendChild(actions);
        tabsList.appendChild(row);
      });
    };

    const renderTags = (tags) => {
      if (!tagsList) return;
      tagsList.innerHTML = '';

      if (!tags.length) {
        const empty = document.createElement('div');
        empty.className = 'empty-state';
        empty.textContent = 'No related tags.';
        tagsList.appendChild(empty);
        return;
      }

      tags.forEach(tag => {
        const row = document.createElement('div');
        row.className = 'assistant-row';

        const title = document.createElement('div');
        title.className = 'assistant-title';
        title.textContent = tag.topic || 'Untitled';

        const meta = document.createElement('div');
        meta.className = 'assistant-meta';
        const domain = (() => {
          try {
            return new URL(tag.url).hostname;
          } catch {
            return 'unknown';
          }
        })();
        meta.textContent = `${this.getTimeAgo(tag.timestamp)} • ${domain}`;

        row.appendChild(title);
        row.appendChild(meta);
        row.onclick = () => {
          if (typeof jumpToHighlight === 'function') {
            jumpToHighlight(tag.id);
          }
        };

        tagsList.appendChild(row);
      });
    };

    const refreshAssistant = async () => {
      if (!tabsList || !tagsList) return;

      if (typeof holoDiagLog === 'function') {
        holoDiagLog('assistant.refresh_start');
      }

      if (!window.chromeTabController) {
        tabsList.textContent = 'Tab access unavailable in this context.';
        tagsList.textContent = 'Related tags unavailable.';
        if (typeof holoDiagLog === 'function') {
          holoDiagLog('assistant.refresh_done', { tabs: 0, tags: 0, error: 'tab_access_unavailable' });
        }
        return;
      }

      tabsList.textContent = 'Loading tabs...';
      tagsList.textContent = 'Loading tags...';

      const [tabs, activeTab, activity] = await Promise.all([
        ChromeTabController.getAllTabs(),
        ChromeTabController.getActiveTab(),
        ChromeTabController.getActivitySnapshot()
      ]);

      const activityMap = activity || {};
      const sortedTabs = tabs
        .slice()
        .sort((a, b) => {
          const aActivity = activityMap[String(a.id)] || {};
          const bActivity = activityMap[String(b.id)] || {};
          const aScore = Math.max(a.lastAccessed || 0, aActivity.lastClick || 0);
          const bScore = Math.max(b.lastAccessed || 0, bActivity.lastClick || 0);
          return bScore - aScore;
        })
        .slice(0, 8);

      renderTabs(sortedTabs, activityMap, activeTab?.id ?? null);

      const activeHost = (() => {
        try {
          return activeTab?.url ? new URL(activeTab.url).hostname : '';
        } catch {
          return '';
        }
      })();

      safeStorageGet({ tags: [] }, (res) => {
        const tags = (res.tags || [])
          .filter(tag => {
            if (!activeHost) return false;
            try {
              return new URL(tag.url).hostname === activeHost;
            } catch {
              return false;
            }
          })
          .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))
          .slice(0, 6);
        renderTags(tags);
        if (typeof holoDiagLog === 'function') {
          holoDiagLog('assistant.refresh_done', { tabs: sortedTabs.length, tags: tags.length });
        }
      });
    };

    if (toggleBtn) {
      toggleBtn.onclick = () => {
        setAssistantEnabled(!assistantEnabledCache, (enabled) => {
          updateToggle(enabled);
          if (typeof holoDiagLog === 'function') {
            holoDiagLog('assistant.toggle', { enabled: enabled });
          }
        });
      };
    }

    if (refreshBtn) {
      refreshBtn.onclick = () => refreshAssistant();
    }

    if (centerBtn) {
      centerBtn.onclick = () => centerPanel();
    }

    loadAssistantEnabled((enabled) => {
      updateToggle(enabled);
      refreshAssistant();
    });
  },
  
  // Create a tag element
  createTagElement: function(tag, isFavorite = false) {
    const tagRow = document.createElement('div');
    tagRow.className = 'tag-row';
    tagRow.dataset.tagId = tag.id;
    
    const favicon = document.createElement('img');
    favicon.className = 'tag-favicon';
    favicon.src = `https://www.google.com/s2/favicons?domain=${new URL(tag.url).hostname}`;
    favicon.onerror = () => { favicon.style.display = 'none'; };
    
    const content = document.createElement('div');
    content.className = 'tag-content';
    
    const label = document.createElement('div');
    label.className = 'tag-label';
    label.textContent = tag.topic || 'Untitled';
    
    const meta = document.createElement('div');
    meta.className = 'tag-meta';
    
    const timeAgo = this.getTimeAgo(tag.timestamp);
    const domain = new URL(tag.url).hostname;
    meta.textContent = `${timeAgo} • ${domain}`;
    
    content.appendChild(label);
    content.appendChild(meta);
    
    const actions = document.createElement('div');
    actions.className = 'tag-actions';
    
    // Add favorite button
    const favoriteBtn = document.createElement('span');
    favoriteBtn.className = 'tag-favorite';
    favoriteBtn.textContent = isFavorite ? '⭐' : '☆';
    favoriteBtn.title = isFavorite ? 'Remove from favorites' : 'Add to favorites';
    favoriteBtn.onclick = (e) => {
      e.stopPropagation();
      this.toggleFavorite(tag.id);
    };
    
    const deleteBtn = document.createElement('span');
    deleteBtn.className = 'tag-delete';
    deleteBtn.textContent = '×';
    deleteBtn.title = 'Delete tag';
    deleteBtn.onclick = (e) => {
      e.stopPropagation();
      this.deleteTag(tag.id);
    };
    
    actions.appendChild(favoriteBtn);
    actions.appendChild(deleteBtn);
    
    tagRow.appendChild(favicon);
    tagRow.appendChild(content);
    tagRow.appendChild(actions);
    
    // Add click handler to jump to tag
    tagRow.onclick = () => {
      if (typeof jumpToHighlight === 'function') {
        jumpToHighlight(tag.id);
      }
    };
    
    return tagRow;
  },
  
  // Helper functions
  getTimeAgo: function(timestamp) {
    if (!timestamp) return 'Unknown';
    const diff = Date.now() - timestamp;
    const minutes = Math.floor(diff / 60000);
    const hours = Math.floor(diff / 3600000);
    const days = Math.floor(diff / 86400000);
    
    if (days > 0) return `${days}d ago`;
    if (hours > 0) return `${hours}h ago`;
    if (minutes > 0) return `${minutes}m ago`;
    return 'Just now';
  },
  
  toggleFavorite: function(tagId) {
    safeStorageGet({ favorites: [] }, (res) => {
      let favorites = res.favorites || [];
      if (favorites.includes(tagId)) {
        favorites = favorites.filter(id => id !== tagId);
      } else {
        favorites.push(tagId);
      }
      
      safeStorageSet({ favorites }, () => {
        console.log(`[TabInterface] Toggled favorite for tag ${tagId}`);
        this.refreshTabs();
      });
    });
  },
  
  deleteTag: function(tagId) {
    if (confirm('Are you sure you want to delete this tag?')) {
      safeStorageGet({ tags: [] }, (res) => {
        const tags = res.tags.filter(tag => tag.id !== tagId);
        safeStorageSet({ tags }, () => {
          console.log(`[TabInterface] Deleted tag ${tagId}`);
          this.refreshTabs();
        });
      });
    }
  },
  
  exportTags: function() {
    safeStorageGet({ tags: [] }, (res) => {
      const tags = res.tags || [];
      const dataStr = JSON.stringify(tags, null, 2);
      const blob = new Blob([dataStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = url;
      a.download = `holotagg-export-${new Date().toISOString().split('T')[0]}.json`;
      a.click();
      
      URL.revokeObjectURL(url);
      console.log('[TabInterface] Tags exported');
    });
  },
  
  clearAllTags: function() {
    if (confirm('Are you sure you want to delete ALL tags? This cannot be undone.')) {
      safeStorageSet({ tags: [], favorites: [] }, () => {
        console.log('[TabInterface] All tags cleared');
        this.refreshTabs();
      });
    }
  },
  
  // Refresh current tab content
  refreshTabs: function() {
    this.loadTabContent(currentActiveTab);
  },
  
  // Get current active tab
  getCurrentTab: function() {
    return currentActiveTab;
  },
  
  // Get all tabs data
  getAllTabs: function() {
    return { ...tabsData };
  },
  
  // Set tab visibility
  setTabVisibility: function(tabId, visible) {
    if (tabsData[tabId]) {
      tabsData[tabId].visible = visible;
      this.refreshTabs();
    }
  },
  
  // Set tab disabled state
  setTabDisabled: function(tabId, disabled) {
    if (tabsData[tabId]) {
      tabsData[tabId].disabled = disabled;
      this.refreshTabs();
    }
  }
};

// Expose tab interface globally for external control
window.holorunTabInterface = TabInterface;

/* ==================== ORGANIZER MANAGER ==================== */

const OrganizerManager = {
  STORAGE_KEYS: {
    items: 'organizerItems',
    folders: 'organizerFolders'
  },
  DEFAULT_FOLDER: 'Inbox',
  MAX_TEXT_LENGTH: 220,

  isInHoloUI: function(element) {
    if (!element) return false;
    const root = element.getRootNode && element.getRootNode();
    if (root && root.host && (root.host.id === 'tag-panel-host' || root.host.id === 'tag-fab-host')) {
      return true;
    }
    if (element.closest) {
      return Boolean(element.closest('#tag-panel-host') || element.closest('#tag-fab-host'));
    }
    return false;
  },

  getSelectedElement: function() {
    const selection = window.getSelection && window.getSelection();
    if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
      const range = selection.getRangeAt(0);
      const node = range.commonAncestorContainer;
      const element = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
      if (element && !this.isInHoloUI(element)) {
        return element;
      }
    }

    if (lastPointerElement && !this.isInHoloUI(lastPointerElement)) {
      return lastPointerElement;
    }

    const active = document.activeElement;
    if (active && !this.isInHoloUI(active)) {
      return active;
    }

    return null;
  },

  getElementSelector: function(element) {
    if (!element || element.nodeType !== 1) return null;
    if (element.id) {
      try {
        return `#${CSS.escape(element.id)}`;
      } catch (_) {
        return `#${element.id.replace(/[^a-zA-Z0-9_-]/g, '')}`;
      }
    }

    const parts = [];
    let current = element;
    let depth = 0;

    while (current && current.nodeType === 1 && depth < 5 && current !== document.body) {
      let selector = current.tagName.toLowerCase();
      const className = (current.className || '').toString().trim();
      if (className) {
        const firstClass = className.split(/\s+/)[0];
        if (firstClass) selector += `.${firstClass.replace(/[^a-zA-Z0-9_-]/g, '')}`;
      }

      const parent = current.parentElement;
      if (parent) {
        const siblings = Array.from(parent.children).filter(child => child.tagName === current.tagName);
        if (siblings.length > 1) {
          const index = siblings.indexOf(current) + 1;
          selector += `:nth-of-type(${index})`;
        }
      }

      parts.unshift(selector);
      current = current.parentElement;
      depth += 1;
    }

    return parts.join(' > ');
  },

  buildElementSnapshot: function(element) {
    const text = (element.innerText || element.textContent || '').trim();
    const snippet = text.length > this.MAX_TEXT_LENGTH ? `${text.slice(0, this.MAX_TEXT_LENGTH)}…` : text;
    const id = element.id ? element.id.trim() : '';
    const className = element.className ? element.className.toString().trim() : '';

    return {
      id: `org_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      url: cleanUrl(location.href),
      pageTitle: document.title || '',
      tagName: element.tagName ? element.tagName.toLowerCase() : 'element',
      elementId: id,
      elementClass: className,
      text: snippet,
      selector: this.getElementSelector(element),
      timestamp: Date.now()
    };
  },

  normalizeFolderName: function(name) {
    const cleaned = (name || '').toString().trim();
    return cleaned || this.DEFAULT_FOLDER;
  },

  ensureFolder: function(folderName, callback) {
    const normalized = this.normalizeFolderName(folderName);
    safeStorageGet({ [this.STORAGE_KEYS.folders]: [this.DEFAULT_FOLDER] }, (res) => {
      const folders = Array.isArray(res[this.STORAGE_KEYS.folders]) ? res[this.STORAGE_KEYS.folders] : [this.DEFAULT_FOLDER];
      if (!folders.includes(normalized)) {
        folders.push(normalized);
        safeStorageSet({ [this.STORAGE_KEYS.folders]: folders }, () => callback(normalized, folders));
      } else {
        callback(normalized, folders);
      }
    });
  },

  captureSelectedElement: function(folderName, callback) {
    const element = this.getSelectedElement();
    if (!element) {
      alert('Select an element or highlight text on the page, then try again.');
      return;
    }

    const snapshot = this.buildElementSnapshot(element);
    this.ensureFolder(folderName, (normalizedFolder) => {
      snapshot.folder = normalizedFolder;
      safeStorageGet({ [this.STORAGE_KEYS.items]: [] }, (res) => {
        const items = Array.isArray(res[this.STORAGE_KEYS.items]) ? res[this.STORAGE_KEYS.items] : [];
        items.unshift(snapshot);
        safeStorageSet({ [this.STORAGE_KEYS.items]: items }, () => {
          if (typeof callback === 'function') callback(snapshot, items);
        });
      });
    });
  },

  updateItemFolder: function(itemId, folderName, callback) {
    const normalized = this.normalizeFolderName(folderName);
    safeStorageGet({ [this.STORAGE_KEYS.items]: [] }, (res) => {
      const items = Array.isArray(res[this.STORAGE_KEYS.items]) ? res[this.STORAGE_KEYS.items] : [];
      const updatedItems = items.map(item => {
        if (item.id === itemId) {
          return { ...item, folder: normalized };
        }
        return item;
      });
      safeStorageSet({ [this.STORAGE_KEYS.items]: updatedItems }, () => {
        this.ensureFolder(normalized, () => {
          if (typeof callback === 'function') callback(updatedItems);
        });
      });
    });
  },

  deleteItem: function(itemId, callback) {
    safeStorageGet({ [this.STORAGE_KEYS.items]: [] }, (res) => {
      const items = Array.isArray(res[this.STORAGE_KEYS.items]) ? res[this.STORAGE_KEYS.items] : [];
      const updatedItems = items.filter(item => item.id !== itemId);
      safeStorageSet({ [this.STORAGE_KEYS.items]: updatedItems }, () => {
        if (typeof callback === 'function') callback(updatedItems);
      });
    });
  },

  jumpToItem: function(item) {
    if (!item || !item.selector) {
      return false;
    }

    if (cleanUrl(location.href) !== cleanUrl(item.url)) {
      window.open(item.url, '_blank');
      return true;
    }

    try {
      const element = document.querySelector(item.selector);
      if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        return true;
      }
    } catch (e) {
      console.warn('[Organizer] Failed to jump to element:', e);
    }

    return false;
  },

  renderOrganizerContent: function(container) {
    container.innerHTML = '';

    const wrapper = document.createElement('div');
    wrapper.className = 'organizer-wrapper';

    const controls = document.createElement('div');
    controls.className = 'organizer-controls';

    const folderInput = document.createElement('input');
    folderInput.type = 'text';
    folderInput.placeholder = 'Folder name (optional)';
    folderInput.className = 'organizer-input';

    const folderSelect = document.createElement('select');
    folderSelect.className = 'organizer-select';

    const captureBtn = document.createElement('button');
    captureBtn.className = 'organizer-button primary';
    captureBtn.textContent = 'Capture Selection';

    const addFolderBtn = document.createElement('button');
    addFolderBtn.className = 'organizer-button';
    addFolderBtn.textContent = 'Add Folder';

    controls.appendChild(folderSelect);
    controls.appendChild(folderInput);
    controls.appendChild(addFolderBtn);
    controls.appendChild(captureBtn);

    const listContainer = document.createElement('div');
    listContainer.className = 'organizer-list';

    wrapper.appendChild(controls);
    wrapper.appendChild(listContainer);
    container.appendChild(wrapper);

    const renderFolders = (folders) => {
      folderSelect.innerHTML = '';
      folders.forEach(folder => {
        const option = document.createElement('option');
        option.value = folder;
        option.textContent = folder;
        folderSelect.appendChild(option);
      });
    };

    const refreshList = () => {
      safeStorageGet({ [this.STORAGE_KEYS.items]: [], [this.STORAGE_KEYS.folders]: [this.DEFAULT_FOLDER] }, (res) => {
        const items = Array.isArray(res[this.STORAGE_KEYS.items]) ? res[this.STORAGE_KEYS.items] : [];
        const folders = Array.isArray(res[this.STORAGE_KEYS.folders]) ? res[this.STORAGE_KEYS.folders] : [this.DEFAULT_FOLDER];
        renderFolders(folders);
        this.renderOrganizerList(listContainer, items, folders);
      });
    };

    captureBtn.onclick = () => {
      const folderName = folderInput.value || folderSelect.value || this.DEFAULT_FOLDER;
      this.captureSelectedElement(folderName, () => {
        folderInput.value = '';
        refreshList();
      });
    };

    addFolderBtn.onclick = () => {
      const folderName = folderInput.value || '';
      if (!folderName.trim()) {
        alert('Enter a folder name first.');
        return;
      }
      this.ensureFolder(folderName, () => {
        folderInput.value = '';
        refreshList();
      });
    };

    refreshList();
  },

  renderOrganizerList: function(container, items, folders) {
    container.innerHTML = '';

    if (!items.length) {
      const empty = document.createElement('div');
      empty.className = 'organizer-empty';
      empty.textContent = 'No captured elements yet. Select something on the page and click “Capture Selection”.';
      container.appendChild(empty);
      return;
    }

    const groups = {};
    items.forEach(item => {
      const folder = this.normalizeFolderName(item.folder || this.DEFAULT_FOLDER);
      if (!groups[folder]) groups[folder] = [];
      groups[folder].push(item);
    });

    Object.keys(groups).forEach(folder => {
      const section = document.createElement('div');
      section.className = 'organizer-section';

      const header = document.createElement('div');
      header.className = 'organizer-section-header';
      header.textContent = `${folder} (${groups[folder].length})`;
      section.appendChild(header);

      groups[folder].forEach(item => {
        const row = this.createOrganizerRow(item, folders);
        section.appendChild(row);
      });

      container.appendChild(section);
    });
  },

  createOrganizerRow: function(item, folders) {
    const row = document.createElement('div');
    row.className = 'organizer-item';

    const title = document.createElement('div');
    title.className = 'organizer-title';
    title.textContent = item.text || `${item.tagName}${item.elementId ? `#${item.elementId}` : ''}`;

    const meta = document.createElement('div');
    meta.className = 'organizer-meta';
    const domain = (() => {
      try {
        return new URL(item.url).hostname;
      } catch (_) {
        return item.url || '';
      }
    })();
    const timeAgo = TabInterface.getTimeAgo(item.timestamp);
    meta.textContent = `${timeAgo} • ${domain}`;

    const info = document.createElement('div');
    info.className = 'organizer-info';
    info.appendChild(title);
    info.appendChild(meta);

    const actions = document.createElement('div');
    actions.className = 'organizer-actions';

    const folderSelect = document.createElement('select');
    folderSelect.className = 'organizer-select small';
    folders.forEach(folder => {
      const option = document.createElement('option');
      option.value = folder;
      option.textContent = folder;
      if ((item.folder || this.DEFAULT_FOLDER) === folder) option.selected = true;
      folderSelect.appendChild(option);
    });
    folderSelect.onchange = () => {
      this.updateItemFolder(item.id, folderSelect.value, () => {
        TabInterface.refreshTabs();
      });
    };

    const openBtn = document.createElement('button');
    openBtn.className = 'organizer-button';
    openBtn.textContent = 'Open';
    openBtn.onclick = () => window.open(item.url, '_blank');

    const jumpBtn = document.createElement('button');
    jumpBtn.className = 'organizer-button';
    jumpBtn.textContent = 'Jump';
    jumpBtn.onclick = () => this.jumpToItem(item);

    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'organizer-button danger';
    deleteBtn.textContent = 'Delete';
    deleteBtn.onclick = () => {
      if (confirm('Delete this captured item?')) {
        this.deleteItem(item.id, () => {
          TabInterface.refreshTabs();
        });
      }
    };

    actions.appendChild(folderSelect);
    actions.appendChild(openBtn);
    actions.appendChild(jumpBtn);
    actions.appendChild(deleteBtn);

    row.appendChild(info);
    row.appendChild(actions);

    return row;
  }
};

// Expose organizer globally
window.holorunOrganizer = OrganizerManager;

/* ==================== CHROME TAB CONTROL INTERFACE ==================== */

// Chrome browser tab control functions
const ChromeTabController = {
  
  // Helper function to send messages to background script
  sendTabMessage: function(type, data = {}) {
    return new Promise((resolve, reject) => {
      if (!chrome.runtime) {
        reject(new Error('Chrome runtime not available'));
        return;
      }
      
      chrome.runtime.sendMessage({ type, ...data }, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        
        if (response && response.success) {
          resolve(response.result);
        } else {
          reject(new Error(response?.error || 'Tab operation failed'));
        }
      });
    });
  },
  
  // Get all Chrome tabs
  getAllTabs: async function() {
    try {
      const tabs = await this.sendTabMessage('TAB_GET_ALL');
      console.log('[ChromeTabController] Retrieved all tabs:', tabs.length);
      return tabs;
    } catch (error) {
      console.error('[ChromeTabController] Error getting all tabs:', error);
      return [];
    }
  },
  
  // Get current active tab
  getActiveTab: async function() {
    try {
      const tab = await this.sendTabMessage('TAB_GET_ACTIVE');
      console.log('[ChromeTabController] Active tab:', tab?.title);
      return tab;
    } catch (error) {
      console.error('[ChromeTabController] Error getting active tab:', error);
      return null;
    }
  },

  getActivitySnapshot: async function() {
    try {
      const snapshot = await new Promise((resolve, reject) => {
        if (!chrome.runtime) {
          reject(new Error('Chrome runtime not available'));
          return;
        }
        chrome.runtime.sendMessage({ type: 'TAB_ACTIVITY_SNAPSHOT' }, (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          resolve(response?.data || {});
        });
      });
      return snapshot || {};
    } catch (error) {
      console.error('[ChromeTabController] Error getting activity snapshot:', error);
      return {};
    }
  },
  
  // Open new tab
  openTab: async function(url, options = {}) {
    try {
      const tab = await this.sendTabMessage('TAB_CREATE', { url, options });
      console.log(`[ChromeTabController] Opened new tab: ${url}`);
      return tab;
    } catch (error) {
      console.error('[ChromeTabController] Error opening tab:', error);
      return null;
    }
  },
  
  // Close tab
  closeTab: async function(tabId) {
    try {
      const result = await this.sendTabMessage('TAB_CLOSE', { tabId });
      console.log(`[ChromeTabController] Closed tab: ${tabId}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error closing tab:', error);
      return false;
    }
  },
  
  // Switch to specific tab
  switchToTab: async function(tabId) {
    try {
      const result = await this.sendTabMessage('TAB_SWITCH', { tabId });
      console.log(`[ChromeTabController] Switched to tab: ${tabId}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error switching tab:', error);
      return false;
    }
  },
  
  // Duplicate current or specific tab
  duplicateTab: async function(tabId = null) {
    try {
      const tab = await this.sendTabMessage('TAB_DUPLICATE', { tabId });
      console.log(`[ChromeTabController] Duplicated tab: ${tabId || 'current'}`);
      return tab;
    } catch (error) {
      console.error('[ChromeTabController] Error duplicating tab:', error);
      return null;
    }
  },
  
  // Pin or unpin tab
  pinTab: async function(tabId = null, pinned = true) {
    try {
      const result = await this.sendTabMessage('TAB_TOGGLE_PIN', { tabId, pinned });
      console.log(`[ChromeTabController] ${pinned ? 'Pinned' : 'Unpinned'} tab: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error toggling pin:', error);
      return false;
    }
  },
  
  // Mute or unmute tab
  muteTab: async function(tabId = null, muted = true) {
    try {
      const result = await this.sendTabMessage('TAB_TOGGLE_MUTE', { tabId, muted });
      console.log(`[ChromeTabController] ${muted ? 'Muted' : 'Unmuted'} tab: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error toggling mute:', error);
      return false;
    }
  },
  
  // Move tab to specific position
  moveTabToPosition: async function(tabId, index) {
    try {
      const result = await this.sendTabMessage('TAB_MOVE', { tabId, index });
      console.log(`[ChromeTabController] Moved tab ${tabId} to position ${index}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error moving tab:', error);
      return false;
    }
  },
  
  // Group tabs together
  groupTabs: async function(tabIds, groupOptions = {}) {
    try {
      const groupId = await this.sendTabMessage('TAB_GROUP', { tabIds, groupOptions });
      console.log(`[ChromeTabController] Grouped tabs:`, tabIds);
      return groupId;
    } catch (error) {
      console.error('[ChromeTabController] Error grouping tabs:', error);
      return null;
    }
  },
  
  // Ungroup tabs
  ungroupTabs: async function(tabIds) {
    try {
      const result = await this.sendTabMessage('TAB_UNGROUP', { tabIds });
      console.log(`[ChromeTabController] Ungrouped tabs:`, tabIds);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error ungrouping tabs:', error);
      return false;
    }
  },
  
  // Close tabs to the right
  closeTabsToRight: async function(tabId = null) {
    try {
      const result = await this.sendTabMessage('TAB_CLOSE_TO_RIGHT', { tabId });
      console.log(`[ChromeTabController] Closed tabs to the right of: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error closing tabs to right:', error);
      return false;
    }
  },
  
  // Close all other tabs
  closeOtherTabs: async function(tabId = null) {
    try {
      const result = await this.sendTabMessage('TAB_CLOSE_OTHERS', { tabId });
      console.log(`[ChromeTabController] Closed other tabs except: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error closing other tabs:', error);
      return false;
    }
  },
  
  // Reload tab
  reloadTab: async function(tabId = null, bypassCache = false) {
    try {
      const result = await this.sendTabMessage('TAB_RELOAD', { tabId, bypassCache });
      console.log(`[ChromeTabController] Reloaded tab: ${tabId || 'current'} (bypass cache: ${bypassCache})`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error reloading tab:', error);
      return false;
    }
  },
  
  // Go back in history
  goBack: async function(tabId = null) {
    try {
      const result = await this.sendTabMessage('TAB_GO_BACK', { tabId });
      console.log(`[ChromeTabController] Went back in tab: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error going back:', error);
      return false;
    }
  },
  
  // Go forward in history
  goForward: async function(tabId = null) {
    try {
      const result = await this.sendTabMessage('TAB_GO_FORWARD', { tabId });
      console.log(`[ChromeTabController] Went forward in tab: ${tabId || 'current'}`);
      return result;
    } catch (error) {
      console.error('[ChromeTabController] Error going forward:', error);
      return false;
    }
  },
  
  // Search tabs by title or URL
  searchTabs: async function(query) {
    try {
      const tabs = await this.sendTabMessage('TAB_SEARCH', { query });
      console.log(`[ChromeTabController] Found ${tabs.length} tabs matching: "${query}"`);
      return tabs;
    } catch (error) {
      console.error('[ChromeTabController] Error searching tabs:', error);
      return [];
    }
  },
  
  // Utility functions
  utils: {
    // Open multiple URLs as tabs
    openMultipleTabs: async function(urls, options = {}) {
      try {
        const tabs = [];
        for (const url of urls) {
          const tab = await ChromeTabController.openTab(url, { 
            active: false, 
            ...options 
          });
          if (tab) tabs.push(tab);
        }
        console.log(`[ChromeTabController] Opened ${tabs.length} tabs`);
        return tabs;
      } catch (error) {
        console.error('[ChromeTabController] Error opening multiple tabs:', error);
        return [];
      }
    },
    
    // Close duplicate tabs
    closeDuplicates: async function() {
      try {
        const tabs = await ChromeTabController.getAllTabs();
        const urlMap = new Map();
        const duplicates = [];
        
        tabs.forEach(tab => {
          if (urlMap.has(tab.url)) {
            duplicates.push(tab.id);
          } else {
            urlMap.set(tab.url, tab.id);
          }
        });
        
        if (duplicates.length > 0) {
          await ChromeTabController.closeTab(duplicates);
          console.log(`[ChromeTabController] Closed ${duplicates.length} duplicate tabs`);
        }
        
        return duplicates.length;
      } catch (error) {
        console.error('[ChromeTabController] Error closing duplicates:', error);
        return 0;
      }
    },
    
    // Get tabs by domain
    getTabsByDomain: async function(domain) {
      try {
        const tabs = await ChromeTabController.getAllTabs();
        const domainTabs = tabs.filter(tab => {
          try {
            const url = new URL(tab.url);
            return url.hostname.includes(domain);
          } catch {
            return false;
          }
        });
        
        console.log(`[ChromeTabController] Found ${domainTabs.length} tabs for domain: ${domain}`);
        return domainTabs;
      } catch (error) {
        console.error('[ChromeTabController] Error getting tabs by domain:', error);
        return [];
      }
    }
  }
};

// Expose Chrome tab controller globally
window.chromeTabController = ChromeTabController;

/* ==================== SCREEN & ELEMENT DIMENSIONS UTILITY ==================== */

// Screen and element dimension utilities
const DimensionUtils = {
  
  // Get screen dimensions (full screen resolution)
  getScreenDimensions: function() {
    return {
      width: screen.width,
      height: screen.height,
      availWidth: screen.availWidth,    // Available width (minus taskbars)
      availHeight: screen.availHeight,  // Available height (minus taskbars)
      colorDepth: screen.colorDepth,
      pixelDepth: screen.pixelDepth,
      orientation: screen.orientation ? screen.orientation.angle : 0
    };
  },
  
  // Get viewport dimensions (browser window visible area)
  getViewportDimensions: function() {
    return {
      width: window.innerWidth,
      height: window.innerHeight,
      outerWidth: window.outerWidth,    // Browser window width including toolbars
      outerHeight: window.outerHeight,  // Browser window height including toolbars
      scrollX: window.scrollX || window.pageXOffset,
      scrollY: window.scrollY || window.pageYOffset,
      devicePixelRatio: window.devicePixelRatio || 1
    };
  },
  
  // Get document dimensions (entire page content)
  getDocumentDimensions: function() {
    const body = document.body;
    const html = document.documentElement;
    
    return {
      width: Math.max(
        body.scrollWidth, body.offsetWidth,
        html.clientWidth, html.scrollWidth, html.offsetWidth
      ),
      height: Math.max(
        body.scrollHeight, body.offsetHeight,
        html.clientHeight, html.scrollHeight, html.offsetHeight
      ),
      clientWidth: html.clientWidth,
      clientHeight: html.clientHeight
    };
  },
  
  // Get element dimensions and position
  getElementDimensions: function(element) {
    if (!element) {
      console.error('[DimensionUtils] Element not provided');
      return null;
    }
    
    const rect = element.getBoundingClientRect();
    const computedStyle = window.getComputedStyle(element);
    
    return {
      // Position relative to viewport
      position: {
        top: rect.top,
        left: rect.left,
        bottom: rect.bottom,
        right: rect.right,
        x: rect.x,
        y: rect.y
      },
      
      // Dimensions
      dimensions: {
        width: rect.width,
        height: rect.height,
        offsetWidth: element.offsetWidth,
        offsetHeight: element.offsetHeight,
        clientWidth: element.clientWidth,
        clientHeight: element.clientHeight,
        scrollWidth: element.scrollWidth,
        scrollHeight: element.scrollHeight
      },
      
      // Position relative to document
      absolutePosition: {
        top: rect.top + window.scrollY,
        left: rect.left + window.scrollX,
        bottom: rect.bottom + window.scrollY,
        right: rect.right + window.scrollX
      },
      
      // CSS computed values
      css: {
        position: computedStyle.position,
        display: computedStyle.display,
        visibility: computedStyle.visibility,
        opacity: computedStyle.opacity,
        zIndex: computedStyle.zIndex,
        transform: computedStyle.transform,
        margin: {
          top: parseInt(computedStyle.marginTop) || 0,
          right: parseInt(computedStyle.marginRight) || 0,
          bottom: parseInt(computedStyle.marginBottom) || 0,
          left: parseInt(computedStyle.marginLeft) || 0
        },
        padding: {
          top: parseInt(computedStyle.paddingTop) || 0,
          right: parseInt(computedStyle.paddingRight) || 0,
          bottom: parseInt(computedStyle.paddingBottom) || 0,
          left: parseInt(computedStyle.paddingLeft) || 0
        },
        border: {
          top: parseInt(computedStyle.borderTopWidth) || 0,
          right: parseInt(computedStyle.borderRightWidth) || 0,
          bottom: parseInt(computedStyle.borderBottomWidth) || 0,
          left: parseInt(computedStyle.borderLeftWidth) || 0
        }
      },
      
      // Visibility checks
      visibility: {
        isVisible: rect.width > 0 && rect.height > 0 && computedStyle.visibility !== 'hidden',
        isInViewport: rect.top >= 0 && rect.left >= 0 && 
                     rect.bottom <= window.innerHeight && rect.right <= window.innerWidth,
        isPartiallyInViewport: rect.bottom > 0 && rect.right > 0 && 
                              rect.top < window.innerHeight && rect.left < window.innerWidth
      }
    };
  },
  
  // Get element by selector with dimensions
  getElementWithDimensions: function(selector) {
    const element = document.querySelector(selector);
    if (!element) {
      console.warn(`[DimensionUtils] Element not found: ${selector}`);
      return null;
    }
    
    return {
      element: element,
      ...this.getElementDimensions(element)
    };
  },
  
  // Get all elements matching selector with their dimensions
  getAllElementsWithDimensions: function(selector) {
    const elements = document.querySelectorAll(selector);
    return Array.from(elements).map(element => ({
      element: element,
      ...this.getElementDimensions(element)
    }));
  },
  
  // Find elements at specific coordinates
  getElementsAtPoint: function(x, y) {
    const elements = document.elementsFromPoint(x, y);
    return elements.map(element => ({
      element: element,
      tagName: element.tagName,
      className: element.className,
      id: element.id,
      ...this.getElementDimensions(element)
    }));
  },
  
  // Calculate center point of element
  getElementCenter: function(element) {
    const rect = element.getBoundingClientRect();
    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
      absoluteX: rect.left + rect.width / 2 + window.scrollX,
      absoluteY: rect.top + rect.height / 2 + window.scrollY
    };
  },
  
  // Calculate distance between two elements
  getElementDistance: function(element1, element2) {
    const center1 = this.getElementCenter(element1);
    const center2 = this.getElementCenter(element2);
    
    const dx = center2.x - center1.x;
    const dy = center2.y - center1.y;
    
    return {
      horizontal: dx,
      vertical: dy,
      diagonal: Math.sqrt(dx * dx + dy * dy),
      angle: Math.atan2(dy, dx) * 180 / Math.PI
    };
  },
  
  // Get viewport quadrants information
  getViewportQuadrants: function() {
    const vp = this.getViewportDimensions();
    const centerX = vp.width / 2;
    const centerY = vp.height / 2;
    
    return {
      center: { x: centerX, y: centerY },
      quadrants: {
        topLeft: { x: 0, y: 0, width: centerX, height: centerY },
        topRight: { x: centerX, y: 0, width: centerX, height: centerY },
        bottomLeft: { x: 0, y: centerY, width: centerX, height: centerY },
        bottomRight: { x: centerX, y: centerY, width: centerX, height: centerY }
      }
    };
  },
  
  // Check if element fits in viewport
  checkElementFit: function(element, padding = 0) {
    const rect = element.getBoundingClientRect();
    const vp = this.getViewportDimensions();
    
    return {
      fitsHorizontally: rect.width + padding * 2 <= vp.width,
      fitsVertically: rect.height + padding * 2 <= vp.height,
      fitsCompletely: rect.width + padding * 2 <= vp.width && rect.height + padding * 2 <= vp.height,
      overflowX: Math.max(0, (rect.width + padding * 2) - vp.width),
      overflowY: Math.max(0, (rect.height + padding * 2) - vp.height)
    };
  },
  
  // Get optimal position for element placement
  getOptimalPosition: function(width, height, preferredPosition = 'center', padding = 20) {
    const vp = this.getViewportDimensions();
    const positions = {};
    
    // Calculate all possible positions
    positions.topLeft = { x: padding, y: padding };
    positions.topRight = { x: vp.width - width - padding, y: padding };
    positions.bottomLeft = { x: padding, y: vp.height - height - padding };
    positions.bottomRight = { x: vp.width - width - padding, y: vp.height - height - padding };
    positions.center = { 
      x: (vp.width - width) / 2, 
      y: (vp.height - height) / 2 
    };
    positions.topCenter = { 
      x: (vp.width - width) / 2, 
      y: padding 
    };
    positions.bottomCenter = { 
      x: (vp.width - width) / 2, 
      y: vp.height - height - padding 
    };
    positions.leftCenter = { 
      x: padding, 
      y: (vp.height - height) / 2 
    };
    positions.rightCenter = { 
      x: vp.width - width - padding, 
      y: (vp.height - height) / 2 
    };
    
    // Check which positions are valid (element fits completely)
    const validPositions = {};
    Object.keys(positions).forEach(key => {
      const pos = positions[key];
      if (pos.x >= 0 && pos.y >= 0 && 
          pos.x + width <= vp.width && pos.y + height <= vp.height) {
        validPositions[key] = pos;
      }
    });
    
    return {
      preferred: positions[preferredPosition] || positions.center,
      allPositions: positions,
      validPositions: validPositions,
      canFit: Object.keys(validPositions).length > 0,
      recommended: validPositions[preferredPosition] || 
                   validPositions.center || 
                   Object.values(validPositions)[0] || 
                   positions.center
    };
  },
  
  // Monitor dimension changes
  createDimensionWatcher: function(callback, throttleMs = 250) {
    let resizeTimer;
    let lastDimensions = this.getViewportDimensions();
    
    const throttledCallback = () => {
      const currentDimensions = this.getViewportDimensions();
      if (currentDimensions.width !== lastDimensions.width || 
          currentDimensions.height !== lastDimensions.height) {
        callback(currentDimensions, lastDimensions);
        lastDimensions = currentDimensions;
      }
    };
    
    const resizeHandler = () => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(throttledCallback, throttleMs);
    };
    
    window.addEventListener('resize', resizeHandler);
    
    // Return cleanup function
    return () => {
      window.removeEventListener('resize', resizeHandler);
      clearTimeout(resizeTimer);
    };
  },
  
  // Comprehensive dimension report
  getAllDimensions: function() {
    return {
      timestamp: Date.now(),
      screen: this.getScreenDimensions(),
      viewport: this.getViewportDimensions(),
      document: this.getDocumentDimensions(),
      quadrants: this.getViewportQuadrants(),
      userAgent: navigator.userAgent,
      platform: navigator.platform,
      language: navigator.language
    };
  },
  
  // Pretty print dimensions for debugging
  logDimensions: function(target = 'all') {
    const dimensions = this.getAllDimensions();
    
    console.group('🖥️ Screen & Element Dimensions');
    
    if (target === 'all' || target === 'screen') {
      console.group('📺 Screen Dimensions');
      console.table(dimensions.screen);
      console.groupEnd();
    }
    
    if (target === 'all' || target === 'viewport') {
      console.group('🪟 Viewport Dimensions');
      console.table(dimensions.viewport);
      console.groupEnd();
    }
    
    if (target === 'all' || target === 'document') {
      console.group('📄 Document Dimensions');
      console.table(dimensions.document);
      console.groupEnd();
    }
    
    console.groupEnd();
    return dimensions;
  }
};

// Expose dimension utilities globally
window.dimensionUtils = DimensionUtils;

function renderPanel() {
  console.log("Rendering panel...");
  if (typeof holoDiagLog === 'function') {
    holoDiagLog('panel.render');
  }

  const getPanelHostCss = (displayValue) => (
    'position: fixed !important; ' +
    'left: 50% !important; ' +
    'top: 50% !important; ' +
    'transform: translate(-50%, -50%) !important; ' +
    'z-index: 2147483647 !important; ' +
    `display: ${displayValue} !important; ` +
    'width: clamp(320px, 58vw, 720px) !important; ' +
    'max-width: calc(100vw - 24px) !important; ' +
    'max-height: min(80vh, calc(100vh - 24px)) !important;'
  );

  const fabStorageKey = 'holorunFabPosition';
  const fabSize = 56;
  const fabMargin = 12;

  function getFabPosition(callback) {
    try {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        chrome.storage.local.get([fabStorageKey], (res) => {
          callback(res ? res[fabStorageKey] : null);
        });
        return;
      }
    } catch {
      // Ignore and fallback.
    }

    try {
      callback(JSON.parse(localStorage.getItem(fabStorageKey) || 'null'));
    } catch {
      callback(null);
    }
  }

  function setFabPosition(value) {
    try {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        chrome.storage.local.set({ [fabStorageKey]: value });
        return;
      }
    } catch {
      // Ignore and fallback.
    }

    try {
      localStorage.setItem(fabStorageKey, JSON.stringify(value));
    } catch {
      // Ignore storage errors.
    }
  }

  function clampFabPosition(left, top) {
    const maxLeft = Math.max(fabMargin, window.innerWidth - fabSize - fabMargin);
    const maxTop = Math.max(fabMargin, window.innerHeight - fabSize - fabMargin);
    return {
      left: Math.min(Math.max(left, fabMargin), maxLeft),
      top: Math.min(Math.max(top, fabMargin), maxTop)
    };
  }

  function applyFabPosition(host, left, top) {
    const clamped = clampFabPosition(left, top);
    host.style.setProperty('left', `${clamped.left}px`, 'important');
    host.style.setProperty('top', `${clamped.top}px`, 'important');
    host.style.setProperty('transform', 'none', 'important');
    return clamped;
  }

  // Create FAB button first (always visible, toggles panel)
  let fabHost = document.getElementById("tag-fab-host");
  if (!fabHost) {
    fabHost = document.createElement("div");
    fabHost.id = "tag-fab-host";
    fabHost.style.cssText = 'position: fixed !important; left: 50% !important; top: 20px !important; transform: translateX(-50%) !important; z-index: 2147483646 !important;';
    document.body.appendChild(fabHost);
    
    const fabShadow = fabHost.attachShadow({ mode: 'open' });
    const fabStyle = document.createElement('style');
    fabStyle.textContent = `:host { all: initial; } #tag-fab { all: initial; display: flex !important; align-items: center !important; justify-content: center !important; width: 56px !important; height: 56px !important; border-radius: 50% !important; background: transparent !important; color: #fff !important; font-size: 28px !important; cursor: pointer !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25) !important; transition: all 0.2s !important; border: 2px solid rgba(17, 24, 39, 0.35) !important; padding: 0 !important; pointer-events: auto !important; } #tag-fab img { width: 34px !important; height: 34px !important; object-fit: contain !important; display: block !important; background: transparent !important; pointer-events: none !important; } #tag-fab:hover { transform: scale(1.1) !important; box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3) !important; } #tag-fab:active { transform: scale(0.95) !important; }`;
    fabShadow.appendChild(fabStyle);
    
    const fab = document.createElement('button');
    fab.id = 'tag-fab';
    const fabLogoUrl = (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getURL)
      ? chrome.runtime.getURL('icon-img/logo.png')
      : 'icon-img/logo.png';
    fab.style.backgroundImage = 'none';
    fab.style.backgroundColor = 'transparent';
    const fabLogo = document.createElement('img');
    fabLogo.src = fabLogoUrl;
    fabLogo.alt = 'Holorun';
    fabLogo.onerror = () => {
      fab.textContent = 'H';
      fab.style.background = 'linear-gradient(135deg, #111827 0%, #374151 100%)';
      fab.style.color = '#fff';
    };
    fab.appendChild(fabLogo);
    fab.title = 'Toggle Tags Panel';
    fab.onclick = () => {
      if (fabHost.dataset.suppressClick === '1') {
        fabHost.dataset.suppressClick = '0';
        return;
      }
      const panelHost = document.getElementById("tag-panel-host");
      if (!panelHost) return;
      const isHidden = panelHost.style.display === 'none' || !panelHost.style.display;
      if (isHidden) {
        panelHost.style.cssText = getPanelHostCss('block');
      } else {
        panelHost.style.cssText = getPanelHostCss('none');
      }
      console.log(`[FAB Click] Panel ${isHidden ? 'shown' : 'hidden'}`);
    };
    fabShadow.appendChild(fab);

    if (!fabHost.dataset.dragInit) {
      fabHost.dataset.dragInit = '1';

      getFabPosition((saved) => {
        if (saved && typeof saved.left === 'number' && typeof saved.top === 'number') {
          applyFabPosition(fabHost, saved.left, saved.top);
        }
      });

      let dragStartX = 0;
      let dragStartY = 0;
      let startLeft = 0;
      let startTop = 0;
      let didMove = false;

      const onPointerMove = (e) => {
        const dx = e.clientX - dragStartX;
        const dy = e.clientY - dragStartY;
        if (!didMove && Math.abs(dx) + Math.abs(dy) < 4) return;
        didMove = true;
        const next = applyFabPosition(fabHost, startLeft + dx, startTop + dy);
        setFabPosition(next);
        e.preventDefault();
      };

      const onPointerUp = () => {
        document.removeEventListener('pointermove', onPointerMove, true);
        if (didMove) {
          fabHost.dataset.suppressClick = '1';
        }
      };

      const onMouseMove = (e) => {
        const dx = e.clientX - dragStartX;
        const dy = e.clientY - dragStartY;
        if (!didMove && Math.abs(dx) + Math.abs(dy) < 4) return;
        didMove = true;
        const next = applyFabPosition(fabHost, startLeft + dx, startTop + dy);
        setFabPosition(next);
        e.preventDefault();
      };

      const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove, true);
        if (didMove) {
          fabHost.dataset.suppressClick = '1';
        }
      };

      const beginDrag = (clientX, clientY) => {
        const rect = fabHost.getBoundingClientRect();
        dragStartX = clientX;
        dragStartY = clientY;
        startLeft = rect.left;
        startTop = rect.top;
        didMove = false;
      };

      fab.addEventListener('pointerdown', (e) => {
        if (e.button !== 0) return;
        beginDrag(e.clientX, e.clientY);
        try {
          if (fab.setPointerCapture) {
            fab.setPointerCapture(e.pointerId);
          }
        } catch {
          // Ignore pointer capture errors.
        }
        document.addEventListener('pointermove', onPointerMove, true);
        document.addEventListener('pointerup', onPointerUp, { once: true, capture: true });
        e.preventDefault();
      }, { passive: false });

      fab.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        beginDrag(e.clientX, e.clientY);
        document.addEventListener('mousemove', onMouseMove, true);
        document.addEventListener('mouseup', onMouseUp, { once: true, capture: true });
        e.preventDefault();
      }, { passive: false });

      fab.addEventListener('dragstart', (e) => {
        e.preventDefault();
      });

      fab.style.touchAction = 'none';
      fab.style.userSelect = 'none';

      window.addEventListener('resize', () => {
        const rect = fabHost.getBoundingClientRect();
        applyFabPosition(fabHost, rect.left, rect.top);
      });
    }
    
    console.log("FAB button created");
  }

  // Try to render even if context check fails - panel rendering is DOM-only
  
  let panelHost = document.getElementById("tag-panel-host");
  let container;
  if (!panelHost) {
    // Create host element for Shadow DOM
    panelHost = document.createElement("div");
    panelHost.id = "tag-panel-host";
    panelHost.style.cssText = getPanelHostCss('none');
    document.body.appendChild(panelHost);
    
    // Attach Shadow DOM to prevent CSS interference
    const shadowRoot = panelHost.attachShadow({ mode: 'open' });
    
    // Inject complete styles into Shadow DOM (smart-ui.css content)
    const styleSheet = document.createElement('style');
    styleSheet.textContent = `
      /* Shadow DOM Panel Styles - Full smart-ui.css injection */
      :host {
        all: initial;
        position: relative !important;
        display: flex !important;
        flex-direction: column !important;
        overflow: hidden !important;
        border-radius: 14px !important;
        border: 1px solid rgba(75, 211, 197, 0.35) !important;
        background: linear-gradient(160deg, #001f2a 0%, #032733 38%, #042f3a 100%) !important;
        box-shadow: 0 20px 55px rgba(0, 0, 0, 0.55), 0 0 0 1px rgba(0, 255, 214, 0.08) inset !important;
      }
      
      * {
        font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif !important;
        box-sizing: border-box !important;
        line-height: 1.4 !important;
      }

      #tag-panel {
        position: absolute !important;
        inset: 0 !important;
        z-index: -1 !important;
        pointer-events: none !important;
        background:
          linear-gradient(180deg, rgba(160, 245, 237, 0.16) 0px, rgba(160, 245, 237, 0.16) 1px, rgba(0, 0, 0, 0) 1px, rgba(0, 0, 0, 0) 100%),
          radial-gradient(circle at 18% 14%, rgba(83, 212, 198, 0.18) 0%, rgba(83, 212, 198, 0) 40%),
          radial-gradient(circle at 82% 10%, rgba(42, 164, 155, 0.2) 0%, rgba(42, 164, 155, 0) 44%),
          linear-gradient(160deg, #001f2a 0%, #032733 38%, #042f3a 100%) !important;
      }
      
      /* Header Styles */
      #tag-panel-header {
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        padding: 16px 20px !important;
        background: linear-gradient(135deg, #111827 0%, #1f2937 100%) !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1) !important;
        gap: 12px !important;
      }
      
      #tag-panel-header span {
        color: #fff !important;
        font-size: 14px !important;
        font-weight: 600 !important;
      }
      
      #tab-count {
        color: rgba(255, 255, 255, 0.6) !important;
        font-size: 12px !important;
        font-weight: 400 !important;
        margin-left: auto !important;
      }
      
      #tag-close-btn {
        background: rgba(255, 255, 255, 0.1) !important;
        color: #fff !important;
        border: none !important;
        width: 32px !important;
        height: 32px !important;
        border-radius: 6px !important;
        cursor: pointer !important;
        font-size: 20px !important;
        line-height: 1 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        transition: all 0.2s ease !important;
        flex-shrink: 0 !important;
      }
      
      #tag-close-btn:hover {
        background: rgba(255, 255, 255, 0.2) !important;
        transform: scale(1.05) !important;
      }
      
      #tag-close-btn:active {
        transform: scale(0.95) !important;
      }
      
      /* Tab Navigation */
      #tab-navigation {
        display: flex !important;
        gap: 4px !important;
        padding: 12px 16px !important;
        background: #1f2937 !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1) !important;
        overflow-x: auto !important;
        scrollbar-width: thin !important;
      }
      
      #tab-navigation::-webkit-scrollbar {
        height: 4px !important;
      }
      
      #tab-navigation::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.2) !important;
        border-radius: 2px !important;
      }
      
      .tab-button {
        display: flex !important;
        align-items: center !important;
        gap: 6px !important;
        padding: 8px 12px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 6px !important;
        color: rgba(255, 255, 255, 0.7) !important;
        font-size: 13px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        white-space: nowrap !important;
        flex-shrink: 0 !important;
      }
      
      .tab-button:hover {
        background: rgba(255, 255, 255, 0.1) !important;
        border-color: rgba(255, 224, 102, 0.3) !important;
        color: #fff !important;
      }
      
      .tab-button.active {
        background: linear-gradient(135deg, rgba(255, 224, 102, 0.15) 0%, rgba(255, 224, 102, 0.08) 100%) !important;
        border-color: #ffe066 !important;
        color: #ffe066 !important;
        font-weight: 600 !important;
      }
      
      .tab-button.disabled {
        opacity: 0.4 !important;
        cursor: not-allowed !important;
      }
      
      .tab-button.disabled:hover {
        background: rgba(255, 255, 255, 0.05) !important;
        border-color: rgba(255, 255, 255, 0.1) !important;
      }
      
      /* Tab Content Area */
      #tab-content {
        flex: 1 !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        background: #111827 !important;
        color: #fff !important;
        padding: 16px !important;
      }
      
      #tab-content::-webkit-scrollbar {
        width: 8px !important;
      }
      
      #tab-content::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.2) !important;
        border-radius: 4px !important;
      }
      
      /* Tags List */
      .tags-list {
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
      }
      
      .tag-row {
        display: flex !important;
        align-items: center !important;
        gap: 10px !important;
        padding: 12px 14px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 8px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        position: relative !important;
      }
      
      .tag-row:hover {
        background: rgba(255, 255, 255, 0.08) !important;
        border-color: rgba(255, 224, 102, 0.3) !important;
        transform: translateX(-4px) !important;
      }
      
      .tag-label {
        flex: 1 !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #fff !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
      }
      
      .tag-meta {
        font-size: 12px !important;
        color: rgba(255, 255, 255, 0.5) !important;
        flex-shrink: 0 !important;
      }
      
      .tag-delete {
        color: #ff5252 !important;
        opacity: 0 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        cursor: pointer !important;
        padding: 4px 6px !important;
        border-radius: 4px !important;
        transition: all 0.15s ease !important;
        flex-shrink: 0 !important;
        background: transparent !important;
        border: none !important;
      }
      
      .tag-row:hover .tag-delete {
        opacity: 0.8 !important;
      }
      
      .tag-delete:hover {
        opacity: 1 !important;
        background: rgba(255, 82, 82, 0.15) !important;
      }
      
      /* Empty State */
      .empty-state {
        text-align: center !important;
        padding: 40px 20px !important;
        color: rgba(255, 255, 255, 0.4) !important;
        font-size: 13px !important;
        line-height: 1.6 !important;
      }
      
      /* Settings Content */
      .settings-content {
        color: #fff !important;
      }
      
      .setting-group {
        margin-bottom: 24px !important;
      }
      
      .setting-group h3 {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #ffe066 !important;
        margin: 0 0 12px 0 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
      }
      
      .setting-group p {
        font-size: 13px !important;
        color: rgba(255, 255, 255, 0.6) !important;
        margin: 4px 0 !important;
        line-height: 1.5 !important;
      }
      
      .setting-item {
        display: flex !important;
        align-items: center !important;
        gap: 8px !important;
        padding: 8px 0 !important;
        color: rgba(255, 255, 255, 0.8) !important;
        font-size: 13px !important;
        cursor: pointer !important;
      }
      
      .setting-item input[type="checkbox"] {
        width: 16px !important;
        height: 16px !important;
        cursor: pointer !important;
      }
      
      .setting-button {
        padding: 10px 16px !important;
        background: rgba(255, 255, 255, 0.1) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 6px !important;
        color: #fff !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        margin-right: 8px !important;
        margin-bottom: 8px !important;
      }
      
      .setting-button:hover {
        background: rgba(255, 255, 255, 0.15) !important;
        border-color: rgba(255, 224, 102, 0.4) !important;
      }
      
      .setting-button.primary {
        background: linear-gradient(135deg, rgba(255, 224, 102, 0.2) 0%, rgba(255, 224, 102, 0.1) 100%) !important;
        border-color: #ffe066 !important;
        color: #ffe066 !important;
      }
      
      .setting-button.primary:hover {
        background: linear-gradient(135deg, rgba(255, 224, 102, 0.3) 0%, rgba(255, 224, 102, 0.15) 100%) !important;
      }
      
      .setting-button.danger {
        background: rgba(255, 82, 82, 0.1) !important;
        border-color: rgba(255, 82, 82, 0.3) !important;
        color: #ff5252 !important;
      }
      
      .setting-button.danger:hover {
        background: rgba(255, 82, 82, 0.2) !important;
        border-color: #ff5252 !important;
      }
      
      /* Workflow Styles */
      .workflow-controls {
        display: flex !important;
        gap: 8px !important;
        margin-bottom: 12px !important;
        flex-wrap: wrap !important;
      }
      
      .workflow-actions {
        display: flex !important;
        gap: 8px !important;
        margin-top: 12px !important;
        flex-wrap: wrap !important;
      }
      
      .workflow-stats {
        background: rgba(255, 255, 255, 0.05) !important;
        padding: 12px !important;
        border-radius: 6px !important;
        font-size: 13px !important;
        color: rgba(255, 255, 255, 0.8) !important;
        line-height: 1.6 !important;
        margin-bottom: 12px !important;
      }
      
      .workflow-meta {
        margin-top: 4px !important;
        font-size: 11px !important;
        color: rgba(255, 255, 255, 0.6) !important;
      }
      
      .workflow-note {
        font-size: 11px !important;
        color: rgba(255, 255, 255, 0.4) !important;
        font-style: italic !important;
        margin-top: 8px !important;
      }
      
      .workflow-grid {
        display: grid !important;
        grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)) !important;
        gap: 12px !important;
        padding: 0 !important;
        width: 100% !important;
      }
      
      .workflow-grid-side { 
        grid-column: 2 !important; 
        grid-row: 1 / span 2 !important; 
      }
      
      .workflow-grid-mini { 
        grid-row: 3 !important; 
      }
      
      .workflow-window-label {
        grid-column: 1 / -1 !important;
        background: linear-gradient(135deg, rgba(255, 224, 102, 0.12) 0%, rgba(255, 224, 102, 0.05) 100%) !important;
        border-left: 3px solid #ffe066 !important;
        padding: 8px 12px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        color: #ffe066 !important;
        border-radius: 4px !important;
        margin-top: 8px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
      }
      
      .workflow-window-label:first-child {
        margin-top: 0 !important;
      }
      
      .workflow-card {
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 8px !important;
        padding: 0 !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        display: flex !important;
        flex-direction: column !important;
        overflow: hidden !important;
        aspect-ratio: 3 / 2 !important;
      }
      
      .workflow-card:hover {
        border-color: rgba(255, 224, 102, 0.3) !important;
        background: rgba(255, 255, 255, 0.08) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
      }
      
      .workflow-card.active {
        border-color: #ffe066 !important;
        background: rgba(255, 224, 102, 0.08) !important;
        box-shadow: 0 0 0 2px rgba(255, 224, 102, 0.2) !important;
      }
      
      .workflow-card.selected {
        border-color: #60a5fa !important;
        background: rgba(96, 165, 250, 0.1) !important;
      }
      
      .workflow-card-chrome {
        display: flex !important;
        align-items: center !important;
        gap: 8px !important;
        padding: 8px 10px !important;
        background: rgba(0, 0, 0, 0.2) !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08) !important;
      }
      
      .workflow-card-favicon {
        width: 16px !important;
        height: 16px !important;
        flex-shrink: 0 !important;
      }
      
      .workflow-card-domain {
        flex: 1 !important;
        font-size: 11px !important;
        color: rgba(255, 255, 255, 0.6) !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
      }
      
      .workflow-card-close {
        width: 20px !important;
        height: 20px !important;
        border: none !important;
        background: rgba(255, 255, 255, 0.1) !important;
        color: #fff !important;
        font-size: 16px !important;
        line-height: 1 !important;
        cursor: pointer !important;
        border-radius: 4px !important;
        flex-shrink: 0 !important;
        transition: all 0.15s ease !important;
        padding: 0 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
      }
      
      .workflow-card-close:hover {
        background: #ff5252 !important;
        transform: scale(1.1) !important;
      }
      
      .workflow-card-preview {
        flex: 1 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        min-height: 80px !important;
        background: linear-gradient(135deg, rgba(100, 150, 255, 0.3) 0%, rgba(150, 100, 255, 0.3) 100%) !important;
      }
      
      .workflow-card-title {
        padding: 0 10px !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        color: #fff !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        margin-top: auto !important;
        padding-top: 8px !important;
      }
      
      .workflow-card-meta {
        padding: 6px 10px 8px !important;
        font-size: 11px !important;
        color: rgba(255, 255, 255, 0.5) !important;
        display: flex !important;
        align-items: center !important;
        gap: 4px !important;
      }
      
      .workflow-card-position {
        font-weight: 600 !important;
        color: #ffe066 !important;
        background: rgba(255, 224, 102, 0.15) !important;
        padding: 2px 6px !important;
        border-radius: 3px !important;
        font-size: 10px !important;
      }
      
      /* Organizer Styles */
      .organizer-input {
        width: 100% !important;
        padding: 10px 12px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 6px !important;
        color: #fff !important;
        font-size: 13px !important;
        margin-bottom: 12px !important;
      }
      
      .organizer-input::placeholder {
        color: rgba(255, 255, 255, 0.4) !important;
      }
      
      .organizer-input:focus {
        outline: none !important;
        border-color: #ffe066 !important;
        background: rgba(255, 255, 255, 0.08) !important;
      }
      
      .organizer-select {
        padding: 10px 12px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 6px !important;
        color: #fff !important;
        font-size: 13px !important;
        cursor: pointer !important;
        margin-bottom: 12px !important;
        width: 100% !important;
      }
      
      .organizer-select.small {
        width: auto !important;
        min-width: 150px !important;
      }
      
      .organizer-select:focus {
        outline: none !important;
        border-color: #ffe066 !important;
      }
      
      /* Assistant Styles */
      .assistant-controls {
        display: flex !important;
        gap: 8px !important;
        margin-bottom: 12px !important;
        flex-wrap: wrap !important;
      }
      
      .assistant-note {
        font-size: 11px !important;
        color: rgba(255, 255, 255, 0.4) !important;
        font-style: italic !important;
        margin-top: 8px !important;
      }
      
      .assistant-list {
        display: flex !important;
        flex-direction: column !important;
        gap: 8px !important;
      }
    `;
    shadowRoot.appendChild(styleSheet);
    
    // Create header with close button
    const header = document.createElement('div');
    header.id = 'tag-panel-header';
    
    // Create logo image (Holorun site icon)
    const logoContainer = document.createElement('div');
    logoContainer.style.cssText = 'display: flex !important; align-items: center !important; gap: 8px !important; flex: 1 !important;';
    logoContainer.innerHTML = `<img src="https://lh3.googleusercontent.com/sitesv/APaQ0SQOcAFLR98nTeycfiORGJ9yCMLx2Re5yVFZrbY3fydrfLdOh3YoI8_TydTuSlwdYVy84PY3fyme48MwAVVz-rHhBqd-8PSczrDK5nQ0zBMhjHus2DyyFVKTV5daD2oDy06RQdxDzHqpxf-_D5dApoBIEBA72zbcQaQLlUT-rC22QXovUFuOuksUO1g=w16383" alt="Holorun" style="width: 24px; height: 24px; flex-shrink: 0; object-fit: contain;" />`;
    
    const titleSpan = document.createElement('span');
    titleSpan.textContent = 'HoloTagg';
    titleSpan.style.cssText = 'font-weight: 600 !important; font-size: 14px !important;';
    
    logoContainer.appendChild(titleSpan);

    const tabCount = document.createElement('span');
    tabCount.id = 'tab-count';
    tabCount.textContent = 'Tabs: --';
    logoContainer.appendChild(tabCount);
    
    const closeBtn = document.createElement('button');
    closeBtn.id = 'tag-close-btn';
    closeBtn.textContent = '✕';
    closeBtn.onclick = () => {
      const panelHost = document.getElementById("tag-panel-host");
      if (!panelHost) return;
      panelHost.style.cssText = getPanelHostCss('none');
      console.log('[Close Button] Panel hidden');
    };
    header.appendChild(logoContainer);
    header.appendChild(closeBtn);
    shadowRoot.appendChild(header);
    
    // Create tab navigation
    const tabNav = document.createElement('div');
    tabNav.id = 'tab-navigation';
    
    // Create tab buttons
    Object.keys(tabsData).forEach(tabId => {
      const tabData = tabsData[tabId];
      if (!tabData.visible) return;
      
      const tabButton = document.createElement('button');
      tabButton.className = 'tab-button';
      tabButton.dataset.tabId = tabId;
      if (tabData.disabled) tabButton.classList.add('disabled');
      if (tabId === currentActiveTab) tabButton.classList.add('active');
      
      const icon = document.createElement('span');
      icon.textContent = tabData.icon;
      
      const label = document.createElement('span');
      label.textContent = tabData.name;
      
      tabButton.appendChild(icon);
      tabButton.appendChild(label);
      
      tabButton.onclick = () => {
        if (!tabData.disabled) {
          TabInterface.switchTab(tabId);
        }
      };
      
      tabNav.appendChild(tabButton);
    });
    
    shadowRoot.appendChild(tabNav);
    
    // Create tab content area
    const tabContent = document.createElement('div');
    tabContent.id = 'tab-content';
    shadowRoot.appendChild(tabContent);
    
    // Create main panel container
    container = document.createElement('div');
    container.id = 'tag-panel';
    shadowRoot.appendChild(container);
    
    // Load initial tab content
    TabInterface.loadTabContent(currentActiveTab);
    
    console.log("Panel created with Shadow DOM isolation (hidden by default)");
  } else {
    // Find existing container in shadow DOM
    container = panelHost.shadowRoot.querySelector('#tag-panel');
    
    // Update tab states and content
    TabInterface.updateTabStates();
    TabInterface.loadTabContent(currentActiveTab);
  }

  updatePanelTabCount();
  if (tabCountIntervalId) {
    clearInterval(tabCountIntervalId);
  }
  tabCountIntervalId = setInterval(updatePanelTabCount, 15000);
  
  // keep reference for other helpers
  panel = container;

  console.log("Panel rendered with tab interface");
}

/* ==================== STORAGE HELPERS ==================== */

/* ==================== SHOW TAG PREVIEW ==================== */

function showTagPreview(tag) {
  const previewContent = document.querySelector('#preview-content');
  if (!previewContent) return;
  
  const domain = new URL(tag.url).hostname;
  const timestamp = new Date(tag.timestamp).toLocaleString();
  
  previewContent.innerHTML = `
    <div style="margin-bottom: 16px !important;">
      <div style="font-size: 14px !important; font-weight: 600 !important; color: #fff !important; margin-bottom: 8px !important;">
        ${tag.topic || 'Tagged Text'}
      </div>
      <div style="font-size: 12px !important; color: rgba(255,255,255,0.7) !important; margin-bottom: 12px !important;">
        📅 ${timestamp}<br>
        🌐 ${domain}
      </div>
    </div>
    
    <div style="background: rgba(255,255,255,0.05) !important; border: 1px solid rgba(255,255,255,0.1) !important; border-radius: 6px !important; padding: 12px !important; margin-bottom: 16px !important;">
      <div style="font-size: 11px !important; color: rgba(255,255,255,0.5) !important; margin-bottom: 6px !important; text-transform: uppercase; letter-spacing: 0.5px;">
        HIGHLIGHTED TEXT
      </div>
      <div style="font-size: 13px !important; line-height: 1.4 !important; color: #fff !important;">
        "${tag.text}"
      </div>
    </div>
    
    ${tag.context && tag.context.before ? `
    <div style="margin-bottom: 12px !important;">
      <div style="font-size: 11px !important; color: rgba(255,255,255,0.5) !important; margin-bottom: 6px !important;">
        CONTEXT
      </div>
      <div style="font-size: 12px !important; color: rgba(255,255,255,0.6) !important; font-style: italic;">
        ...${tag.context.before.slice(-60)} <span style="background: rgba(255,255,0,0.2); padding: 2px 4px; border-radius: 3px;">[highlighted]</span> ${tag.context.after.slice(0, 60)}...
      </div>
    </div>
    ` : ''}
    
    <div style="border-top: 1px solid rgba(255,255,255,0.1) !important; padding-top: 12px !important; margin-top: 16px !important;">
      <div style="font-size: 11px !important; color: rgba(255,255,255,0.5) !important; margin-bottom: 8px !important;">
        ACTIONS
      </div>
      <div style="display: flex !important; gap: 8px !important; flex-wrap: wrap !important;">
        <button style="background: rgba(59,130,246,0.2) !important; border: 1px solid rgba(59,130,246,0.3) !important; color: #60a5fa !important; padding: 6px 12px !important; border-radius: 4px !important; font-size: 11px !important; cursor: pointer;" 
                onclick="window.open('${tag.url}${tag.url.includes('#') ? '&' : '#'}holorunTagId=${encodeURIComponent(tag.id)}', '_blank')">
          🔗 Open Page
        </button>
        ${cleanUrl(location.href) === cleanUrl(tag.url) ? `
        <button style="background: rgba(16,185,129,0.2) !important; border: 1px solid rgba(16,185,129,0.3) !important; color: #34d399 !important; padding: 6px 12px !important; border-radius: 4px !important; font-size: 11px !important; cursor: pointer;"
                onclick="jumpToHighlight('${tag.id}')">
          ⚡ Jump Here
        </button>
        ` : ''}
        <button style="background: rgba(239,68,68,0.2) !important; border: 1px solid rgba(239,68,68,0.3) !important; color: #f87171 !important; padding: 6px 12px !important; border-radius: 4px !important; font-size: 11px !important; cursor: pointer;"
                onclick="if(confirm('Delete this tag?')) deleteTag('${tag.id}')">
          🗑️ Delete
        </button>
      </div>
    </div>
  `;
}

/* ==================== DELETE TAG ==================== */

function deleteTag(tagId) {
  console.log("Deleting:", tagId);
  
  try {
    safeStorageGet({ tags: [] }, (res) => {
    const tags = res.tags || [];
    const newTags = tags.filter(t => t.id !== tagId);
    
    safeStorageSet({ tags: newTags }, () => {
      console.log("✓ Deleted, remaining:", newTags.length);
      removeHighlight(tagId);
      renderPanel();
    });
  });
  } catch (e) {
    console.warn('[DeleteTag] Extension context lost:', e.message);
  }
}

// -------- Cross-node text matching fallback helpers --------
function normalizeSimple(s) {
  return (s || '').replace(/\s+/g, ' ').trim();
}

function collectAnchorText(node, maxLen = 160) {
  try {
    let cursor = node;
    let depth = 0;
    while (cursor && depth < 3) {
      const text = (cursor.innerText || '').trim();
      if (text && text.length > 0) {
        return text.substring(0, maxLen);
      }
      cursor = cursor.parentElement;
      depth++;
    }
  } catch (_) {}
  return '';
}

function extractIconMeta(el, labelOverride = '') {
  if (!el) return null;
  const label = labelOverride || el.getAttribute('aria-label') || el.getAttribute('title') || (el.dataset && el.dataset.icon) || '';
  
  // Calculate position index among similar icons
  const selector = '[aria-label], [data-icon], [role="img"], [class*="icon"], i';
  const allCandidates = Array.from(document.querySelectorAll(selector));
  const matchingIndex = allCandidates.findIndex(candidate => candidate === el);
  
  const meta = {
    label: (label || '').trim(),
    tagName: el.tagName ? el.tagName.toLowerCase() : '',
    classList: el.classList ? Array.from(el.classList) : [],
    ariaLabel: el.getAttribute ? (el.getAttribute('aria-label') || '') : '',
    title: el.getAttribute ? (el.getAttribute('title') || '') : '',
    dataIcon: el.dataset ? (el.dataset.icon || '') : '',
    role: el.getAttribute ? (el.getAttribute('role') || '') : '',
    anchorText: collectAnchorText(el),
    iconIndex: matchingIndex >= 0 ? matchingIndex : null  // Track position for multi-icon restoration
  };
  return meta;
}

function tokenizeWords(s) {
  const cleaned = (s || '').toLowerCase().replace(/[^a-z0-9\s]+/g, ' ');
  return cleaned.split(/\s+/).filter(Boolean);
}

// --- Bullet point helpers ---
function isBulletMarker(ch) {
  return /^(?:[\u2022\u25E6\u25AA\u25CF\-*\u2013\u2014]|[\u2705\u274C\u2714\u2611\u26D4]|\[[ xX]\])$/.test(ch);
}

function looksLikeBulletText(s) {
  const raw = s || '';
  const norm = normalizeTextForMatching(raw);
  // Match common bullet prefixes: •, -, –, *, emoji checks, [x], numeric lists "1.", "2)", "a)"
  const bulletPrefixRe = /^\s*(?:[\u2022\u25E6\u25AA\u25CF\-*\u2013\u2014]|[\u2705\u274C\u2714\u2611\u26D4]|\[[ xX]\]|\d+[\.\)]|[a-zA-Z][\.\)])\s+/;
  const m = norm.match(bulletPrefixRe);
  if (!m) return { isBullet: false, stripped: norm };
  const stripped = norm.replace(bulletPrefixRe, '');
  return { isBullet: true, stripped };
}

function findRangeInListItem(targetText) {
  try {
    const { isBullet, stripped } = looksLikeBulletText(targetText);
    const searchText = normalizeTextForMatching(isBullet ? stripped : targetText);
    if (!searchText) return null;

    const root = getContentRoot();
    const items = Array.from(root.querySelectorAll('li'));
    if (!items.length) {
      console.log('[FindBullet] ❌ No list items found');
      return null;
    }

    let best = null;
    let bestScore = -1;
    items.forEach((li) => {
      const liText = normalizeTextForMatching(li.innerText || '');
      if (!liText) return;
      // Score: direct includes preferred; else token overlap
      let score = 0;
      if (liText.includes(searchText)) score += 5;
      else if (searchText.includes(liText)) score += 3; // very short bullets
      // token overlap
      const tA = tokenizeWords(searchText);
      const tB = tokenizeWords(liText);
      const setB = new Set(tB);
      let overlap = 0;
      for (const t of tA) if (setB.has(t)) overlap++;
      score += Math.min(overlap, 6);

      if (score > bestScore) {
        bestScore = score;
        best = li;
      }
    });

    if (!best || bestScore < 3) {
      console.log('[FindBullet] ❌ No sufficiently similar list item found');
      return null;
    }

    // Build a range within the LI using linearized text to find the exact snippet
    const { full, mapping } = linearizeTextNodes(best);
    const normalizedFull = normalizeTextForMatching(full);
    const escaped = searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'i');
    let match = normalizedFull.match(pattern);
    if (!match) {
      const idx = normalizedFull.indexOf(searchText);
      if (idx >= 0) match = { index: idx, 0: searchText };
    }
    let range;
    if (match) {
      const startIdx = match.index ?? 0;
      const endIdx = startIdx + (match[0] ? match[0].length : searchText.length);
      range = positionToRange(mapping, startIdx, endIdx);
    }
    if (!range) {
      // Fallback: select beginning of LI up to max length
      range = document.createRange();
      // Choose first text node in LI
      const walker = document.createTreeWalker(best, NodeFilter.SHOW_TEXT, null);
      const first = walker.nextNode();
      const lastNode = (() => {
        let n = first; let prev = first;
        while (n) { prev = n; n = walker.nextNode(); }
        return prev || first;
      })();
      if (!first) return null;
      range.setStart(first, 0);
      range.setEnd(lastNode, Math.min((lastNode && (lastNode.nodeValue || '').length) || 0, 500));
    }

    // Clamp to single LI and safe length
    const clamped = clampRangeToSingleBlock(range, 500);
    if (clamped) range = clamped;
    const preview = (range.toString() || '').substring(0, 50);
    console.log(`[FindBullet] ✅ Found in LI - length: ${range.toString().length} chars, preview: "${preview}..."`);
    return range;
  } catch (e) {
    console.warn('[FindBullet] error:', e.message);
    return null;
  }
}

// --- Header helpers ---
function isHeaderTagName(tag) {
  return ['H1','H2','H3','H4','H5','H6'].includes((tag || '').toUpperCase());
}

function findRangeInHeader(targetText) {
  try {
    const searchText = normalizeTextForMatching(targetText);
    const root = getContentRoot();
    const headers = Array.from(root.querySelectorAll('h1,h2,h3,h4,h5,h6'));
    if (!headers.length) return null;

    let best = null;
    let bestScore = -1;
    headers.forEach((h) => {
      const txt = normalizeTextForMatching(h.innerText || '');
      if (!txt) return;
      let score = 0;
      if (txt.includes(searchText)) score += 5;
      const tA = tokenizeWords(searchText);
      const tB = tokenizeWords(txt);
      const setB = new Set(tB);
      let overlap = 0;
      for (const t of tA) if (setB.has(t)) overlap++;
      score += Math.min(overlap, 6);
      if (score > bestScore) { bestScore = score; best = h; }
    });
    if (!best || bestScore < 3) return null;
    const { full, mapping } = linearizeTextNodes(best);
    const escaped = searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const rawPattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'i');
    const rawMatch = full.match(rawPattern);

    let startIdx = -1;
    let endIdx = -1;

    if (rawMatch) {
      startIdx = rawMatch.index ?? 0;
      endIdx = startIdx + (rawMatch[0] ? rawMatch[0].length : searchText.length);
    } else {
      // Fallback: normalized match, then re-align on raw text
      const normalizedFull = normalizeTextForMatching(full);
      const normPattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'i');
      let match = normalizedFull.match(normPattern);
      if (!match) {
        const idx = normalizedFull.indexOf(searchText);
        if (idx >= 0) match = { index: idx, 0: searchText };
      }
      if (!match) return null;
      const normSnippet = match[0] || searchText;
      // Re-find the matched snippet inside raw text to keep offsets accurate
      const rawLower = full.toLowerCase();
      const normLower = normSnippet.toLowerCase();
      const approxIdx = rawLower.indexOf(normLower);
      if (approxIdx >= 0) {
        startIdx = approxIdx;
        endIdx = startIdx + normLower.length;
      } else {
        startIdx = match.index ?? 0;
        endIdx = startIdx + searchText.length;
      }
    }

    if (startIdx < 0 || endIdx <= startIdx) return null;

    const range = positionToRange(mapping, startIdx, endIdx);
    if (!range) return null;
    const clamped = clampRangeToSingleBlock(range, 500);
    return clamped || range;
  } catch (e) {
    console.warn('[FindHeader] error:', e.message);
    return null;
  }
}

// --- Number-aware matching ---
function buildNumberFlexibleRegex(text) {
  const normalized = normalizeTextForMatching(text);
  // Replace digit runs with a flexible \\d+ pattern; leave words intact
  const parts = normalized.split(/(\d+)/);
  const escapedParts = parts.map(p => /\d+/.test(p) ? '\\d+' : p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'));
  const pattern = escapedParts.join('');
  return new RegExp(pattern, 'i');
}

function findRangeByRawTextNumbers(targetText) {
  try {
    const { full, mapping } = linearizeTextNodes(getContentRoot());
    if (!full || !mapping.length) return null;
    const normalizedFull = normalizeTextForMatching(full);
    const re = buildNumberFlexibleRegex(targetText);
    let match = normalizedFull.match(re);
    if (!match) return null;
    const startIdx = match.index ?? normalizedFull.search(re);
    const endIdx = startIdx + (match[0] ? match[0].length : normalizeTextForMatching(targetText).length);
    const range = positionToRange(mapping, startIdx, endIdx);
    if (!range) return null;
    return clampRangeToSingleBlock(range, 500) || range;
  } catch (e) {
    console.warn('[RawTextNumbers] error:', e.message);
    return null;
  }
}

// --- Kind classification ---
function getNearestBlockTag(node) {
  try {
    let el = node && (node.nodeType === 1 ? node : node.parentElement);
    while (el && el !== document.body) {
      if (['LI','P','DIV','SECTION','ARTICLE','H1','H2','H3','H4','H5','H6'].includes(el.tagName)) return el.tagName;
      el = el.parentElement;
    }
  } catch (_) {}
  return '';
}

function classifyTagKind(text, containerTag, iconMeta) {
  const norm = normalizeTextForMatching(text || '');
  const bullet = looksLikeBulletText(norm);
  const hasIcon = !!iconMeta;
  const hasNumber = /\d/.test(norm);
  const isHeader = isHeaderTagName(containerTag);
  const isParagraph = containerTag === 'P' || containerTag === 'DIV' || containerTag === 'SECTION' || containerTag === 'ARTICLE';
  if (hasIcon && bullet.isBullet) return 'bullet+icon';
  if (hasIcon && isHeader && hasNumber) return 'header+number+icon';
  if (hasIcon && isHeader) return 'header+icon';
  if (hasIcon) return 'icon';
  if (bullet.isBullet || containerTag === 'LI') return 'bullet';
  if (isHeader) return hasNumber ? 'header+number' : 'header';
  if (hasNumber) return 'text+numbers';
  if (isParagraph) return 'paragraph';
  return 'text';
}

function findRangeByTokens(targetText, tagId = null) {
  try {
    // Early exit if highlight already exists
    if (tagId && document.querySelector(`[data-tag-id="${tagId}"]`)) {
      return null;
    }
    const tokens = tokenizeWords(targetText);
    if (!tokens.length) return null;
    const slice = tokens.slice(0, Math.min(6, tokens.length));
    if (!slice.length) return null;
    console.log(`[FindRangeTokens] Attempting token fallback for: "${targetText.substring(0, 40)}..."`);


    const root = getContentRoot();
    const host = location.hostname || '';
    const isGmail = /mail\.google\.com$/.test(host);
    
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => {
        if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        const parent = node.parentElement;
        if (!parent || typeof parent.closest !== 'function') return NodeFilter.FILTER_REJECT;
        // Exclude our UI and non-content areas
        if (parent.closest('#tag-popup, #tag-panel, .tag-highlight, script, style, noscript, head')) {
          return NodeFilter.FILTER_REJECT;
        }
        // Exclude SVG, canvas, diagrams, code blocks (React-managed content)
        if (parent.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"], [class*="code"]')) {
          return NodeFilter.FILTER_REJECT;
        }
        // Exclude interactive and editable regions
        if (parent.closest('input, textarea, [contenteditable], [role="textbox"], [role="searchbox"], form, button, [role="button"], label')) {
          return NodeFilter.FILTER_REJECT;
        }
        // Gmail-specific exclusions to reduce false positives
        if (isGmail) {
          // Exclude Gmail navigation, toolbars, and UI elements
          if (parent.closest('.aeN, .aKz, .ar7, .aDP, .bhZ, .G-asx, .nZ, .Cr.aqJ, [role="navigation"], [class*="toolbar"], [class*="sidebar"], [data-tooltip], [aria-label*="navigation"], [aria-label*="toolbar"]')) {
            return NodeFilter.FILTER_REJECT;
          }
          // Exclude Gmail status indicators and notifications
          if (parent.closest('.aT, .ag, .ar, .as, .ao4, [class*="loading"], [class*="status"]')) {
            return NodeFilter.FILTER_REJECT;
          }
        }
        // Exclude sidebars/nav/headers/footers
        if (parent.closest('[class*="history" i], [class*="sidebar" i], nav, aside, [role="navigation"], header, footer, [class*="footer" i], [aria-label*="sidebar" i], [aria-label*="navigation" i], [data-testid*="sidebar" i], [id*="sidebar" i]')) {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    // Simple contiguous text match before cross-node token fallback
    let startNode = null, startOffset = 0;
    let endNode = null, endOffset = 0;
    let currentTokenIndex = 0;
    let searchingStart = true;

    const firstToken = slice[0];
    while (walker.nextNode()) {
      const node = walker.currentNode;
      const raw = node.nodeValue || '';
      const norm = normalizeSimple(raw).toLowerCase();
      if (!norm) continue;

      if (searchingStart) {
        const pos = norm.indexOf(firstToken);
        if (pos !== -1) {
          const rawPos = raw.toLowerCase().indexOf(firstToken);
          if (rawPos !== -1) {
            startNode = node;
            startOffset = rawPos;
            searchingStart = false;
            currentTokenIndex = 1;
            if (slice.length === 1) {
              endNode = node;
              endOffset = rawPos + firstToken.length;
              // Evaluate single-token match immediately
              const candidate = document.createRange();
              candidate.setStart(startNode, Math.max(0, startOffset));
              candidate.setEnd(endNode, Math.max(0, endOffset));

              const candidateText = candidate.toString();
              console.log(`[FindRangeTokens] ✅ Found via token fallback - length: ${candidateText.length} chars, preview: "${candidateText.substring(0, 50)}..."`);

              const MAX_REASONABLE_TOKEN_RANGE = 5000;
              if (candidateText.length > MAX_REASONABLE_TOKEN_RANGE) {
                console.log(`[FindRangeTokens] ❌ Discarding token match; span ${candidateText.length} exceeds ${MAX_REASONABLE_TOKEN_RANGE} chars - likely wrong. Searching for tighter match...`);
                startNode = endNode = null;
                startOffset = endOffset = 0;
                searchingStart = true;
                currentTokenIndex = 0;
                continue;
              }

              // If original text looks like a bullet, prefer LI ancestor
              const bulletCheck = looksLikeBulletText(targetText);
              if (bulletCheck.isBullet) {
                const block = getPreferredBlockAncestor(startNode, 1200);
                if (!block || block.tagName !== 'LI') {
                  console.log('[FindRangeTokens] ⚠️ Bullet text matched outside LI; rejecting candidate');
                  startNode = endNode = null;
                  startOffset = endOffset = 0;
                  searchingStart = true;
                  currentTokenIndex = 0;
                  continue;
                }
              }

              return candidate;
            }
          }
        }
      } else {
        let searchFrom = 0;
        while (currentTokenIndex < slice.length) {
          const token = slice[currentTokenIndex];
          const npos = norm.indexOf(token, searchFrom);
          if (npos === -1) {
            break;
          }
          const rawPos = raw.toLowerCase().indexOf(token, searchFrom);
          if (rawPos === -1) {
            searchFrom = npos + token.length;
            continue;
          }
          endNode = node;
          endOffset = rawPos + token.length;
          currentTokenIndex += 1;
          searchFrom = rawPos + token.length;
        }
        if (currentTokenIndex >= slice.length) {
          // Evaluate the candidate range before accepting
          const candidate = document.createRange();
          candidate.setStart(startNode, Math.max(0, startOffset));
          candidate.setEnd(endNode, Math.max(0, endOffset));

          const candidateText = candidate.toString();
          console.log(`[FindRangeTokens] ✅ Found via token fallback - length: ${candidateText.length} chars, preview: "${candidateText.substring(0, 50)}..."`);

          const MAX_REASONABLE_TOKEN_RANGE = 5000;
          if (candidateText.length > MAX_REASONABLE_TOKEN_RANGE) {
            console.log(`[FindRangeTokens] ❌ Discarding token match; span ${candidateText.length} exceeds ${MAX_REASONABLE_TOKEN_RANGE} chars - likely wrong. Searching for tighter match...`);
            // Reset search and continue walking for a tighter span
            startNode = endNode = null;
            startOffset = endOffset = 0;
            searchingStart = true;
            currentTokenIndex = 0;
            continue;
          }

          // If original text looks like a bullet, prefer LI ancestor
          const bulletCheck2 = looksLikeBulletText(targetText);
          if (bulletCheck2.isBullet) {
            const block = getPreferredBlockAncestor(startNode, 1200);
            if (!block || block.tagName !== 'LI') {
              console.log('[FindRangeTokens] ⚠️ Bullet text matched outside LI; rejecting candidate');
              startNode = endNode = null;
              startOffset = endOffset = 0;
              searchingStart = true;
              currentTokenIndex = 0;
              continue;
            }
          }

          return candidate;
        }
      }
    }
    console.log(`[FindRangeTokens] ❌ Token fallback failed`);
    return null;
  } catch (e) {
    console.warn('[HolorunTagger/CrossNode] findRangeByTokens error:', e);
    return null;
  }
}

function findIconElement(iconMeta) {
  if (!iconMeta) return null;
  try {
    const candidates = Array.from(document.querySelectorAll('[aria-label], [data-icon], [role="img"], [class*="icon"], i'));
    
    // If iconIndex is known, try to find the exact nth matching icon first
    if (iconMeta.iconIndex !== null && iconMeta.iconIndex !== undefined) {
      const exactMatch = candidates[iconMeta.iconIndex];
      if (exactMatch && !exactMatch.closest('#tag-popup, #tag-panel, .tag-highlight, script, style, noscript, head') &&
          !exactMatch.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"]') &&
          !exactMatch.closest('input, textarea, [contenteditable], [role="textbox"], [role="searchbox"], form, button, [role="button"], label')) {
        console.log(`[Icon] Using exact iconIndex match: ${iconMeta.iconIndex}`);
        return exactMatch;
      }
    }
    
    // Fallback: score-based matching
    let best = null;
    let bestScore = -1;
    candidates.forEach((el) => {
      if (!el || !el.getAttribute || !el.closest) return;
      // Exclude our UI and non-content areas
      if (el.closest('#tag-popup, #tag-panel, .tag-highlight, script, style, noscript, head')) return;
      if (el.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"]')) return;
      if (el.closest('input, textarea, [contenteditable], [role="textbox"], [role="searchbox"], form, button, [role="button"], label')) return;

      if (iconMeta.tagName && el.tagName && el.tagName.toLowerCase() !== iconMeta.tagName) return;

      const aria = (el.getAttribute('aria-label') || '').trim();
      const title = (el.getAttribute('title') || '').trim();
      const dataIcon = (el.dataset && el.dataset.icon) || '';
      const classes = el.classList ? Array.from(el.classList) : [];
      let score = 0;

      if (iconMeta.ariaLabel && aria === iconMeta.ariaLabel) score += 3;
      if (iconMeta.title && title === iconMeta.title) score += 2;
      if (iconMeta.dataIcon && dataIcon === iconMeta.dataIcon) score += 2;
      if (iconMeta.label && (aria === iconMeta.label || title === iconMeta.label)) score += 2;
      if (iconMeta.classList && iconMeta.classList.length) {
        const overlap = iconMeta.classList.filter(c => classes.includes(c)).length;
        score += overlap;
      }
      if (iconMeta.anchorText) {
        const localText = collectAnchorText(el).toLowerCase();
        if (localText && localText.includes(iconMeta.anchorText.toLowerCase())) score += 2;
      }

      if (score > bestScore) {
        bestScore = score;
        best = el;
      }
    });
    return best;
  } catch (e) {
    console.warn('[HolorunTagger/Icon] findIconElement error:', e);
    return null;
  }
}

function wrapElementWithHighlight(el, tag) {
  if (!el || !el.parentNode) return null;
  const existing = el.closest('.tag-highlight');
  if (existing) return existing;
  try {
    // Set flag to prevent MutationObserver from triggering during highlight creation
    window.__holorunCreatingHighlight = true;

    const span = document.createElement('span');
    span.className = 'tag-highlight';
    span.dataset.tagId = tag.id;
    span.dataset.tag = tag.topic;
    
    // Apply explicit inline styles with !important to beat host CSS overrides
    const setStyle = (prop, val) => {
      try { span.style.setProperty(prop, val, 'important'); } catch (_) {}
    };
    setStyle('background', 'linear-gradient(135deg, #ffe066 0%, #ffd54f 100%)');
    setStyle('background-color', '#ffe066');
    setStyle('border-radius', '3px');
    setStyle('padding', '0 2px');
    setStyle('cursor', 'pointer');
    setStyle('display', 'inline');
    setStyle('position', 'relative');
    setStyle('line-height', 'inherit');
    setStyle('font-size', 'inherit');
    setStyle('font-family', 'inherit');
    setStyle('font-weight', 'inherit');
    setStyle('letter-spacing', 'inherit');
    setStyle('word-spacing', 'inherit');
    setStyle('mix-blend-mode', 'normal');
    setStyle('isolation', 'isolate');
    setStyle('box-shadow', '0 1px 3px rgba(255, 224, 102, 0.2)');
    setStyle('color', 'inherit');
    
    el.parentNode.insertBefore(span, el);
    span.appendChild(el);

    // Clear flag after a brief delay
    setTimeout(() => {
      window.__holorunCreatingHighlight = false;
    }, 100);

    return span;
  } catch (e) {
    console.warn('[HolorunTagger/Icon] wrapElementWithHighlight error:', e);
    window.__holorunCreatingHighlight = false;
    return null;
  }
}

function pulseHighlight(span) {
  if (!span) return;
  // Set flag to prevent MutationObserver from triggering during pulse styling
  window.__holorunCreatingHighlight = true;

  span.style.transition = 'outline 0.25s ease, box-shadow 0.25s ease, background-color 0.25s ease';
  const prevOutline = span.style.outline;
  const prevOutlineOffset = span.style.outlineOffset;
  const prevBg = span.style.backgroundColor;
  const prevShadow = span.style.boxShadow;
  span.style.outline = '3px solid #ff9800';
  span.style.outlineOffset = '3px';
  span.style.backgroundColor = 'rgba(255,152,0,0.15)';
  span.style.boxShadow = '0 0 0 2px rgba(255,152,0,0.2)';
  setTimeout(() => {
    span.style.outline = prevOutline;
    span.style.outlineOffset = prevOutlineOffset;
    span.style.backgroundColor = prevBg;
    span.style.boxShadow = prevShadow;
    // Clear flag after pulse completes
    window.__holorunCreatingHighlight = false;
  }, 1200);
}

  // --- Helper: Detect nearest scrollable ancestor and sticky top overlays ---
  function getScrollableAncestor(el) {
    try {
      let node = el.parentElement;
      while (node && node !== document.body && node !== document.documentElement) {
        const style = window.getComputedStyle(node);
        const overflowY = style.overflowY;
        const overflow = style.overflow;
        const isScrollable = (overflowY === 'auto' || overflowY === 'scroll' || overflow === 'auto' || overflow === 'scroll');
        const hasScroll = node.scrollHeight > node.clientHeight + 4; // tolerance
        if (isScrollable && hasScroll) return node;
        node = node.parentElement;
      }
    } catch (_) {}
    return null;
  }

  function getFixedTopOverlapHeight() {
    try {
      const elements = Array.from(document.querySelectorAll('body *'));
      let total = 0;
      for (const el of elements) {
        const style = window.getComputedStyle(el);
        if (style.position !== 'fixed') continue;
        const rect = el.getBoundingClientRect();
        // Count fixed bars pinned to the very top that overlap viewport
        if (rect.top <= 0 && rect.bottom > 0 && rect.height > 20 && rect.width > 200) {
          total += rect.height;
        }
      }
      // Cap to a reasonable value
      return Math.min(total, Math.floor(window.innerHeight * 0.4));
    } catch (_) {
      return 0;
    }
  }

  function ensureVisible(el) {
    try {
      const scrollParent = getScrollableAncestor(el);
      const rect = el.getBoundingClientRect();
      const fixedTop = getFixedTopOverlapHeight();
      const targetTopViewport = Math.max(0, rect.top - fixedTop);

      if (scrollParent) {
        const parentRect = scrollParent.getBoundingClientRect();
        const offsetTop = rect.top - parentRect.top + scrollParent.scrollTop - fixedTop;
        scrollParent.scrollTo({ top: offsetTop - Math.max(0, scrollParent.clientHeight / 3), behavior: 'smooth' });
        console.log(`[Scroll] Scrolled parent container (fixedTop=${fixedTop}px)`);
      } else {
        // Window scroll with absolute target so far-offscreen highlights land correctly
        const targetY = Math.max(0, window.scrollY + rect.top - fixedTop - Math.max(0, window.innerHeight / 3));
        window.scrollTo({ top: targetY, behavior: 'smooth' });
        console.log(`[Scroll] Scrolled window (fixedTop=${fixedTop}px, targetY=${targetY})`);
      }
    } catch (e) {
      console.warn('[Scroll] ensureVisible failed:', e.message);
      try { el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' }); } catch (_) {}
    }
  }

function restoreTagHighlight(tagId, opts = { pulse: true, scroll: true }) {
  // Early exit if highlight already exists
  let existing = document.querySelector(`[data-tag-id="${tagId}"]`);
  if (existing) {
    if (opts.scroll) ensureVisible(existing);
    if (opts.pulse) pulseHighlight(existing);
    return;
  }
  
  // TRY to restore even if context check fails - storage API might still work
  
  try {
    safeStorageGet({ tags: [] }, (res) => {
    const tags = res.tags || [];
    const tag = tags.find(t => t.id === tagId);
    if (!tag) {
      console.warn('[Restore] Tag not found in storage:', tagId);
      return;
    }

    const matchText = getMatchableText(tag);
    const kind = tag.kind || classifyTagKind(matchText, tag.containerTag || '', tag.iconMeta);

    let highlight = document.querySelector(`[data-tag-id="${tagId}"]`);
    if (!highlight) {
      if (tag.iconMeta) {
        const iconEl = findIconElement(tag.iconMeta);
        if (iconEl) {
          highlight = wrapElementWithHighlight(iconEl, tag);
        }
      }
    }

    if (!highlight) {
      let range = null;
      // Strategy order based on kind
      switch (kind) {
        case 'icon':
          // icon handled earlier; if not found, no range
          break;
        case 'bullet':
        case 'bullet+icon':
          console.log('[Restore] Kind=bullet; using LI-scoped search');
          range = findRangeInListItem(matchText) || findRangeByRawText(matchText, tag.context);
          break;
        case 'header':
        case 'header+number':
        case 'header+icon':
        case 'header+number+icon':
          console.log('[Restore] Kind=header; searching header elements');
          range = findRangeInHeader(matchText) || findRangeByRawText(matchText, tag.context);
          break;
        case 'text+numbers':
          console.log('[Restore] Kind=text+numbers; using number-flexible match');
          range = findRangeByRawTextNumbers(matchText) || findRangeByRawText(matchText, tag.context);
          break;
        case 'paragraph':
          // Prefer raw match then clamp
          range = findRangeByRawText(matchText, tag.context);
          if (range) {
            const clamped = clampRangeToSingleBlock(range, 500);
            if (clamped) range = clamped;
          }
          break;
        default:
          range = findRangeByRawText(matchText, tag.context);
      }

      // Then tokenized fallback
      if (!range) {
        console.log(`[Restore] Anchor/primary match failed, trying token fallback...`);
        range = findRangeByTokens(matchText, tagId);
      }
      // As final fallback, try loose raw-text without anchors
      if (!range) {
        console.log(`[Restore] Token fallback failed, trying loose raw-text fallback...`);
        range = findRangeByRawTextLoose(matchText);
      }
      if (!range) {
        console.warn(`[Restore] Could not locate text for tag: ${tagId}`);
        return;
      }
      highlight = surroundWithHighlight(range, tag);
    }

    if (!highlight) {
      console.warn(`[Restore] Failed to create highlight for: ${tagId}`);
      return;
    }

    // Cache after successful find
    if (typeof cacheHighlightPosition === 'function') {
      cacheHighlightPosition(tag);
    }

    if (opts.scroll) {
      try {
        highlight.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
      } catch (_) {}
    }
    if (opts.pulse) {
      pulseHighlight(highlight);
    }
    });
  } catch (e) {
    console.warn('[RestoreTagHighlight] Extension context lost:', e.message);
  }
}

// Linearize all text nodes into one string with position mapping
// Position map = array tracking which DOM node each character belongs to
// Example: "Hello world" from 2 nodes → [{node: node1, start:0, end:6}, {node: node2, start:6, end:11}]
// Matching configuration for anchor validation
const MATCH_CONFIG = {
  requireBothAnchors: false,   // relaxed: at least one anchor must match for flexibility on content changes
  maxAnchorLen: 60,            // cap anchors to 60 chars for balanced accuracy and flexibility
  normalizeLevel: 'whitespace' // reserved for future (whitespace-only normalization for now)
};
function linearizeTextNodes(root) {
  try {
    const host = location.hostname || '';
    const isGmail = /mail\.google\.com$/.test(host);
    
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => {
        if (!node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        const parent = node.parentElement;
        if (!parent || typeof parent.closest !== 'function') return NodeFilter.FILTER_REJECT;
        // Exclude our UI and core non-content
        if (parent.closest('#tag-popup, #tag-panel, .tag-highlight, script, style, noscript, head')) return NodeFilter.FILTER_REJECT;
        // Exclude code/diagrams
        if (parent.closest('svg, canvas, pre, code, .mermaid, [class*="diagram"], [class*="code"]')) return NodeFilter.FILTER_REJECT;
        // Exclude forms and inputs
        if (parent.closest('input, textarea, [contenteditable], [role="textbox"], [role="searchbox"], form, button, [role="button"], label')) return NodeFilter.FILTER_REJECT;
        // Gmail-specific exclusions to reduce false positives
        if (isGmail) {
          // Exclude Gmail navigation, toolbars, and UI elements
          if (parent.closest('.aeN, .aKz, .ar7, .aDP, .bhZ, .G-asx, .nZ, .Cr.aqJ, [role="navigation"], [class*="toolbar"], [class*="sidebar"], [data-tooltip], [aria-label*="navigation"], [aria-label*="toolbar"]')) return NodeFilter.FILTER_REJECT;
          // Exclude Gmail status indicators and notifications
          if (parent.closest('.aT, .ag, .ar, .as, .ao4, [class*="loading"], [class*="status"]')) return NodeFilter.FILTER_REJECT;
        }
        // Exclude rendered/history content - BUT ALLOW message containers
        // This is the key fix: message containers like [role="article"] should NOT be rejected
        if (parent.closest('[class*="history"], [class*="sidebar"], nav, [role="navigation"], footer, [class*="footer"]')) return NodeFilter.FILTER_REJECT;
        // Do NOT exclude [role="article"] or message content - ChatGPT messages live there
        // if (parent.closest('[role="article"]')) return NodeFilter.FILTER_REJECT;
        // Exclude dropdowns, tooltips, popups
        if (parent.closest('[role="menu"], [role="listbox"], [role="dialog"], [class*="dropdown"], [class*="tooltip"], [class*="popup"]')) return NodeFilter.FILTER_REJECT;
        // Exclude aria-hidden regions to avoid off-screen/duplicated clones
        if (parent.closest('[aria-hidden="true"]')) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    const parts = [];
    const mapping = [];
    let position = 0;

    while (walker.nextNode()) {
      const node = walker.currentNode;
      const text = node.nodeValue || '';
      if (!text) continue;

      parts.push(text);
      mapping.push({ node, start: position, end: position + text.length });
      position += text.length;
    }

    return { full: parts.join(''), mapping };
  } catch (e) {
    console.warn('[Linearize] Error:', e);
    return { full: '', mapping: [] };
  }
}

// Convert string position back to DOM range using position map
function positionToRange(mapping, startIdx, endIdx) {
  try {
    if (!mapping || !mapping.length) return null;

    // Find which text node(s) contain our start/end positions
    let startNode = null, startOffset = 0;
    let endNode = null, endOffset = 0;

    for (const entry of mapping) {
      // Start position
      if (startIdx >= entry.start && startIdx < entry.end) {
        startNode = entry.node;
        startOffset = startIdx - entry.start;
      }
      // End position
      if (endIdx > entry.start && endIdx <= entry.end) {
        endNode = entry.node;
        endOffset = endIdx - entry.start;
      }
      if (startNode && endNode) break;
    }

    if (!startNode || !endNode) return null;

    const range = document.createRange();
    range.setStart(startNode, startOffset);
    range.setEnd(endNode, endOffset);
    return range;
  } catch (e) {
    console.warn('[PositionToRange] Error:', e);
    return null;
  }
}

// Find range using linearized text + anchor validation (context before/after)
// CRITICAL: First finds EXACT text match, only uses context to disambiguate multiple matches
function findRangeByRawText(targetText, context = null) {
  try {
    const normalizedTarget = normalizeTextForMatching(targetText);
    if (!normalizedTarget) return null;

    // Linearize all text nodes into one string with position map, scoped to content root
    const { full, mapping } = linearizeTextNodes(getContentRoot());
    if (!full || !mapping.length) return null;
    const normalizedFull = normalizeTextForMatching(full);

    // Build whitespace-flexible regex with GLOBAL flag to find ALL matches
    const escaped = normalizedTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'gi');

    // Find ALL occurrences of the exact text
    const allMatches = [];
    let match;
    while ((match = pattern.exec(normalizedFull)) !== null) {
      allMatches.push({
        index: match.index,
        text: match[0],
        startIdx: match.index,
        endIdx: match.index + match[0].length
      });
    }

    console.log(`   🔍 Found ${allMatches.length} exact text match(es)`);

    if (allMatches.length === 0) {
      // Fallback to indexOf for exact match
      const idx = normalizedFull.indexOf(normalizedTarget);
      if (idx >= 0) {
        allMatches.push({
          index: idx,
          text: normalizedTarget,
          startIdx: idx,
          endIdx: idx + normalizedTarget.length
        });
        console.log(`   🔍 Found 1 exact text match via indexOf`);
      } else {
        console.log(`   ❌ No exact text matches found`);
        return null;
      }
    }

    // EXACT MATCH LOGIC:
    // If only ONE match found → use it immediately (most common case)
    if (allMatches.length === 1) {
      console.log(`   ✅ EXACT MATCH: Only one occurrence found, using it directly`);
      const startIdx = allMatches[0].startIdx;
      const endIdx = allMatches[0].endIdx;
      
      const result = positionToRange(mapping, startIdx, endIdx);
      if (result) {
        const resultText = result.toString();
        console.log(`   ✅ Strategy 1 SUCCESS: Exact single match`);
        console.log(`   📏 Match Length: ${resultText.length} chars`);
        console.log(`   📄 Match Preview: "${resultText.substring(0, 80)}${resultText.length > 80 ? '...' : ''}"`);
        console.log(`   📍 Position: string index ${startIdx}-${endIdx}`);
      } else {
        console.log(`   ❌ Strategy 1 FAILED: positionToRange couldn't create DOM range`);
      }
      return result;
    }

    // MULTIPLE MATCHES: Use context to disambiguate
    console.log(`   ⚠️ Multiple matches found (${allMatches.length}), using context to disambiguate...`);
    
    if (!context || (!context.before && !context.after)) {
      console.warn(`   ⚠️ No context available to disambiguate ${allMatches.length} matches, using first match`);
      const startIdx = allMatches[0].startIdx;
      const endIdx = allMatches[0].endIdx;
      
      const result = positionToRange(mapping, startIdx, endIdx);
      if (result) {
        const resultText = result.toString();
        console.log(`   ⚠️ Strategy 1 SUCCESS: First match (no context)`);
        console.log(`   📏 Match Length: ${resultText.length} chars`);
        console.log(`   📄 Match Preview: "${resultText.substring(0, 80)}${resultText.length > 80 ? '...' : ''}"`);
      }
      return result;
    }

    // Check context anchors for each match to find the right one
    const normalizeWS = (s) => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
    const beforeLen = context.before ? (MATCH_CONFIG.maxAnchorLen ? Math.min(context.before.length, MATCH_CONFIG.maxAnchorLen) : context.before.length) : 0;
    const afterLen = context.after ? (MATCH_CONFIG.maxAnchorLen ? Math.min(context.after.length, MATCH_CONFIG.maxAnchorLen) : context.after.length) : 0;
    const beforeAnchorNorm = context.before ? normalizeWS(context.before.slice(Math.max(0, context.before.length - beforeLen))) : '';
    const afterAnchorNorm = context.after ? normalizeWS(context.after.slice(0, afterLen)) : '';
    
    const checkAnchors = (sIdx, eIdx) => {
      const beforeSliceNorm = beforeLen
        ? normalizeWS(full.slice(Math.max(0, sIdx - beforeLen), sIdx))
        : '';
      const afterSliceNorm = afterLen
        ? normalizeWS(full.slice(eIdx, Math.min(full.length, eIdx + afterLen)))
        : '';

      const beforeOk = beforeLen ? beforeSliceNorm.endsWith(beforeAnchorNorm) : true;
      const afterOk = afterLen ? afterSliceNorm.startsWith(afterAnchorNorm) : true;
      return MATCH_CONFIG.requireBothAnchors ? (beforeOk && afterOk) : (beforeOk || afterOk);
    };

    // Find the match with valid anchors
    for (let i = 0; i < allMatches.length; i++) {
      const candidate = allMatches[i];
      if (checkAnchors(candidate.startIdx, candidate.endIdx)) {
        console.log(`   ✅ CONTEXT MATCH: Match #${i + 1} passed anchor validation`);
        
        // CRITICAL: Verify that the matched text at this position actually contains our target
        const matchedTextAtPosition = full.substring(candidate.startIdx, candidate.endIdx);
        const matchedNorm = normalizeTextForMatching(matchedTextAtPosition);
        if (matchedNorm !== normalizedTarget) {
          console.log(`   ❌ VALIDATION FAILED: Text at position doesn't match (got "${matchedTextAtPosition}", expected "${targetText}")`);
          continue;
        }
        
        const result = positionToRange(mapping, candidate.startIdx, candidate.endIdx);
        if (result) {
          const resultText = result.toString();
          console.log(`   ✅ Strategy 1 SUCCESS: Context-disambiguated match`);
          console.log(`   📏 Match Length: ${resultText.length} chars`);
          console.log(`   📄 Match Preview: "${resultText.substring(0, 80)}${resultText.length > 80 ? '...' : ''}"`);
          console.log(`   📍 Position: string index ${candidate.startIdx}-${candidate.endIdx}`);
        } else {
          console.log(`   ❌ Strategy 1 FAILED: positionToRange couldn't create DOM range`);
        }
        return result;
      } else {
        console.log(`   ❌ Match #${i + 1} failed anchor validation`);
      }
    }

    console.warn(`   ❌ No match found with valid anchors (all ${allMatches.length} matches failed anchor check)`);
    return null;
  } catch (e) {
    console.warn('[HolorunTagger/RawMatch] findRangeByRawText error:', e);
    return null;
  }
}

// Loose raw-text search without anchor validation (whitespace-flexible)
function findRangeByRawTextLoose(targetText) {
  try {
    const normalizedTarget = normalizeTextForMatching(targetText);
    if (!normalizedTarget) return null;
    console.log(`[FindRangeLoose] Attempting loose (no-anchor) fallback for: "${normalizedTarget.substring(0, 40)}..."`);

    const { full, mapping } = linearizeTextNodes(getContentRoot());
    if (!full || !mapping.length) {
      console.log(`[FindRangeLoose] ❌ No text nodes found (linearization returned empty)`);
      return null;
    }
    const normalizedFull = normalizeTextForMatching(full);

    const escaped = normalizedTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(escaped.replace(/\s+/g, '\\s+'), 'i');

    let match = normalizedFull.match(pattern);
    if (!match) {
      const idx = normalizedFull.indexOf(normalizedTarget);
      if (idx >= 0) {
        match = { index: idx, 0: normalizedTarget };
      }
    }
    if (!match) {
      console.log(`[FindRangeLoose] ❌ Text not found in linearized DOM (${full.length} chars scanned)`);
      console.log(`[FindRangeLoose]    Searched for: "${normalizedTarget.substring(0, 60)}..."`);
      return null;
    }

    const startIdx = match.index ?? 0;
    const endIdx = startIdx + (match[0] ? match[0].length : targetText.length);
    const result = positionToRange(mapping, startIdx, endIdx);
    
    if (!result) {
      console.log(`[FindRangeLoose] ❌ positionToRange failed (text found at index ${startIdx} but couldn't create range)`);
      return null;
    }
    
    // Log the matched text length
    const resultText = result.toString();
    console.log(`[FindRangeLoose] Found match - length: ${resultText.length} chars, preview: "${resultText.substring(0, 50)}..."`);
    
    // Validate matched text is not unreasonably large
    if (resultText.length > 500) {
      console.log(`[FindRangeLoose] ⚠️ WARNING: Matched text is ${resultText.length} chars (likely over-matched). This may be wrong.`);
    }
    
    // Validate that the matched range is NOT in an irrelevant container
    if (result && result.commonAncestorContainer) {
      const ancestor = result.commonAncestorContainer.parentElement || result.commonAncestorContainer;
      if (ancestor && typeof ancestor.closest === 'function') {
        // Reject if in chat/message/history containers
        if (ancestor.closest('[role="article"], [data-test-id*="message"], .message, .chat-message, [class*="message"]')) {
          console.log(`[FindRangeLoose] ⚠️ Match rejected: found in message container`);
          return null;
        }
        if (ancestor.closest('[class*="history"], [class*="sidebar"], nav, footer')) {
          console.log(`[FindRangeLoose] ⚠️ Match rejected: found in sidebar/history/nav`);
          return null;
        }
      }
    }
    
    if (result) {
      console.log(`[FindRangeLoose] ✅ Found via loose fallback`);
    }
    return result;
  } catch (e) {
    console.warn('[HolorunTagger/RawMatchLoose] error:', e);
    return null;
  }
}

// Trim a DOM range to a maximum text length while preserving the start point
function trimRangeToLength(range, maxLen) {
  try {
    const text = range.toString();
    if (text.length <= maxLen) return range;

    const newRange = document.createRange();
    let remaining = maxLen;
    let started = false;

    const walker = document.createTreeWalker(
      range.commonAncestorContainer,
      NodeFilter.SHOW_TEXT,
      {
        acceptNode: (node) => range.intersectsNode(node)
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT
      }
    );

    while (walker.nextNode()) {
      const node = walker.currentNode;
      const value = node.nodeValue || '';
      let nodeStart = 0;
      let nodeEnd = value.length;

      if (node === range.startContainer) {
        nodeStart = range.startOffset;
      }
      if (node === range.endContainer) {
        nodeEnd = Math.min(nodeEnd, range.endOffset);
      }

      if (nodeEnd <= nodeStart) continue;

      if (!started) {
        newRange.setStart(node, nodeStart);
        started = true;
      }

      const spanLen = nodeEnd - nodeStart;
      if (spanLen >= remaining) {
        newRange.setEnd(node, nodeStart + remaining);
        return newRange;
      }

      remaining -= spanLen;
    }

    // Fallback: if we exhaust nodes, clamp to original end
    newRange.setEnd(range.endContainer, range.endOffset);
    return newRange;
  } catch (e) {
    console.warn('[Highlight/Trim] Failed to trim range:', e.message);
    return null;
  }
}

// Find a sensible block ancestor (prefer list items, else smallest block with manageable text)
function getPreferredBlockAncestor(node, maxTextLen = 1200) {
  try {
    const BLOCK_TAGS = new Set(['LI', 'P', 'DIV', 'SECTION', 'ARTICLE', 'MAIN', 'ASIDE', 'DD', 'DT', 'FIGURE', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TD', 'TH', 'BLOCKQUOTE', 'UL', 'OL']);
    let cursor = node && node.nodeType === 1 ? node : node && node.parentElement;
    let firstBlock = null;
    let firstListItem = null;
    while (cursor && cursor !== document.body && cursor !== document.documentElement) {
      if (BLOCK_TAGS.has(cursor.tagName)) {
        if (!firstBlock) firstBlock = cursor;
        if (cursor.tagName === 'LI' && !firstListItem) firstListItem = cursor;
        const textLen = (cursor.innerText || '').length;
        if (textLen > 0 && textLen <= maxTextLen) {
          return cursor; // close-enough small block
        }
      }
      cursor = cursor.parentElement;
    }
    if (firstListItem) return firstListItem;
    return firstBlock;
  } catch (_) {
    return null;
  }
}

// Trim a range but keep it scoped within a root element (e.g., a single paragraph/list item)
function trimRangeWithinRoot(range, root, maxLen) {
  try {
    if (!root || !root.contains(range.startContainer)) return null;

    const newRange = document.createRange();
    let remaining = maxLen;
    let started = false;

    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) => range.intersectsNode(node) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT
    });

    while (walker.nextNode()) {
      const node = walker.currentNode;
      const value = node.nodeValue || '';
      let nodeStart = 0;
      let nodeEnd = value.length;

      if (node === range.startContainer) {
        nodeStart = Math.max(nodeStart, range.startOffset);
      }
      if (node === range.endContainer) {
        nodeEnd = Math.min(nodeEnd, range.endOffset);
      }

      if (nodeEnd <= nodeStart) continue;

      if (!started) {
        newRange.setStart(node, nodeStart);
        started = true;
      }

      const spanLen = nodeEnd - nodeStart;
      if (spanLen >= remaining) {
        newRange.setEnd(node, nodeStart + remaining);
        return newRange;
      }

      remaining -= spanLen;
    }

    if (started) {
      newRange.setEnd(range.endContainer, Math.min(range.endOffset, (range.endContainer.nodeValue || '').length));
      return newRange;
    }
    return null;
  } catch (e) {
    console.warn('[Highlight/TrimRoot] Failed to trim within root:', e.message);
    return null;
  }
}

// Constrain a range to a single logical block to avoid wrapping entire messages
function clampRangeToSingleBlock(range, maxLen) {
  try {
    const block = getPreferredBlockAncestor(range.startContainer, maxLen * 2);
    if (!block) return range;
    const trimmed = trimRangeWithinRoot(range, block, maxLen);
    if (trimmed && trimmed.toString().length <= maxLen) {
      return trimmed;
    }
    return range;
  } catch (_) {
    return range;
  }
}

function surroundWithHighlight(range, tag) {
  try {
    // CRITICAL: Validate range size before applying - prevent highlighting whole page
    const MAX_REASONABLE_HIGHLIGHT = 500; // chars

    let rangeText = range.toString();
    if (rangeText.length > MAX_REASONABLE_HIGHLIGHT) {
      console.warn(`[Highlight/Validation] ❌ REJECTED: Range too large (${rangeText.length} chars, max ${MAX_REASONABLE_HIGHLIGHT})`);
      console.warn(`[Highlight/Validation] Text preview: "${rangeText.substring(0, 100)}..."`);
      console.warn(`[Highlight/Validation] Attempting to trim to safe length...`);
      const trimmed = trimRangeToLength(range, MAX_REASONABLE_HIGHLIGHT);
      if (!trimmed) {
        console.warn(`[Highlight/Validation] Trimming failed; aborting highlight.`);
        return null;
      }
      range = trimmed;
      rangeText = range.toString();
      console.warn(`[Highlight/Validation] ✅ Trimmed oversized range to ${rangeText.length} chars for safe highlighting`);
    }

    // Constrain range to a single logical block to avoid wrapping whole messages or multiple bullets
    const clamped = clampRangeToSingleBlock(range, MAX_REASONABLE_HIGHLIGHT);
    if (clamped && clamped !== range) {
      range = clamped;
      rangeText = range.toString();
      console.log(`[Highlight/Validation] ✅ Clamped range to single block (${rangeText.length} chars)`);
    }
    
    // Check if range spans across multiple major containers (sign of over-matching)
    const commonAncestor = range.commonAncestorContainer;
    if (commonAncestor.nodeType === 1) { // Element node
      const ancestor = commonAncestor;
      // Reject if ancestor is body or document
      if (ancestor.tagName === 'BODY' || ancestor.tagName === 'HTML' || ancestor === document) {
        console.warn(`[Highlight/Validation] ❌ REJECTED: Range spans entire body/html`);
        return null;
      }
    }
    
    // Set flag to prevent MutationObserver from triggering during highlight creation
    window.__holorunCreatingHighlight = true;
    
    // Create highlight span
    const span = document.createElement('span');
    span.className = 'tag-highlight';
    span.dataset.tagId = tag.id;
    span.dataset.tag = tag.topic;
    
    // Apply explicit inline styles with !important to beat host CSS overrides
    const setStyle = (prop, val) => {
      try { span.style.setProperty(prop, val, 'important'); } catch (_) {}
    };
    setStyle('background', 'linear-gradient(135deg, #ffe066 0%, #ffd54f 100%)');
    setStyle('background-color', '#ffe066');
    setStyle('border-radius', '3px');
    setStyle('padding', '0 2px');
    setStyle('cursor', 'pointer');
    setStyle('display', 'inline');  // CRITICAL: Must be inline, not block
    setStyle('position', 'relative');
    setStyle('line-height', 'inherit');
    setStyle('font-size', 'inherit');
    setStyle('font-family', 'inherit');
    setStyle('font-weight', 'inherit');
    setStyle('letter-spacing', 'inherit');
    setStyle('word-spacing', 'inherit');
    setStyle('mix-blend-mode', 'normal');
    setStyle('isolation', 'isolate');
    setStyle('box-shadow', '0 1px 3px rgba(255, 224, 102, 0.2)');
    setStyle('color', 'inherit');
    setStyle('white-space', 'normal');  // Preserve whitespace handling
    setStyle('text-decoration', 'none');  // No underline
    setStyle('border', 'none');  // No borders
    
    // Log diagnostic info
    const safeRangeText = range.toString();
    const rangeTextPreview = safeRangeText.substring(0, 50);
    console.log(`[Highlight/Create] Range text: "${rangeTextPreview}..."`);
    console.log(`[Highlight/Create] Range length: ${safeRangeText.length} chars`);
    console.log(`[Highlight/Create] Tag: ${tag.topic}`);
    
    // Use surroundContents to wrap the range in the span
    // This preserves the original nodes without duplication
    try {
      range.surroundContents(span);
      console.log(`[Highlight/Create] ✅ Successfully surrounded content with tag-highlight span`);
    } catch (hierarchyErr) {
      // Fallback: if surroundContents fails (e.g., partial node selection in block elements),
      // use a safer DOM-aware wrapping method that preserves layout
      console.log(`[Highlight/Create] ⚠️ surroundContents failed (${hierarchyErr.message}), using safe wrap-nodes method`);
      
      // Safe fallback: walk through range endpoints and wrap text nodes individually
      // This avoids extracting/reinserting which can cause layout shifts
      const walker = document.createTreeWalker(
        range.commonAncestorContainer,
        NodeFilter.SHOW_TEXT,
        null,
        false
      );
      
      const nodesToWrap = [];
      let node;
      const rangeStart = range.startContainer;
      const rangeEnd = range.endContainer;
      const rangeStartOffset = range.startOffset;
      const rangeEndOffset = range.endOffset;
      let isInRange = false;
      
      // Collect text nodes in range
      while (node = walker.nextNode()) {
        if (node === rangeStart) isInRange = true;
        if (isInRange) nodesToWrap.push(node);
        if (node === rangeEnd) break;
      }
      
      if (nodesToWrap.length > 0) {
        // For simple single-node case, just wrap it
        if (nodesToWrap.length === 1) {
          const node = nodesToWrap[0];
          const beforeNode = node.splitText(rangeStartOffset);
          const afterNode = beforeNode.splitText(rangeEndOffset - rangeStartOffset);
          span.appendChild(beforeNode);
          beforeNode.parentNode.insertBefore(span, afterNode);
        } else {
          // Multi-node: use extract but keep in place more carefully
          const contents = range.extractContents();
          span.appendChild(contents);
          range.insertNode(span);
        }
        console.log(`[Highlight/Create] ✅ Fallback safe wrap complete (${nodesToWrap.length} nodes)`);
      } else {
        // Last resort: use extract method
        const contents = range.extractContents();
        span.appendChild(contents);
        range.insertNode(span);
        console.log(`[Highlight/Create] ✅ Fallback extract insertion complete`);
      }
    }
    
    // Log computed styles after creation - DETAILED DIAGNOSTIC
    setTimeout(() => {
      const styles = window.getComputedStyle(span);
      console.log(`[Highlight/CSS] ✅ INLINE STYLES SET:`);
      console.log(`[Highlight/CSS]   - inline background: ${span.style.background}`);
      console.log(`[Highlight/CSS]   - inline backgroundColor: ${span.style.backgroundColor}`);
      console.log(`[Highlight/CSS]   - inline borderRadius: ${span.style.borderRadius}`);
      console.log(`[Highlight/CSS]   - inline padding: ${span.style.padding}`);
      console.log(`[Highlight/CSS] ✅ COMPUTED STYLES:`);
      console.log(`[Highlight/CSS]   - computed background: ${styles.background}`);
      console.log(`[Highlight/CSS]   - computed display: ${styles.display}`);
      console.log(`[Highlight/CSS]   - computed visibility: ${styles.visibility}`);
      console.log(`[Highlight/CSS]   - computed opacity: ${styles.opacity}`);
      console.log(`[Highlight/CSS] ✅ CLASS: ${span.className}`);
      window.__holorunCreatingHighlight = false;
    }, 50);
    
    return span;
  } catch (e) {
    console.warn('[HolorunTagger/CrossNode] surroundWithHighlight error:', e);
    window.__holorunCreatingHighlight = false;
    return null;
  }
}

/* ==================== EXPAND COLLAPSED CHATGPT MESSAGES ==================== */

/**
 * ChatGPT collapse/expand detection
 * 
 * NOTE: Disabled by design - ChatGPT's aria-expanded="false" elements are primarily
 * UI controls (dropdowns, toggles, form inputs), NOT message content toggles.
 * Clicking these breaks the UI and causes extension context invalidation.
 * 
 * Real issue: ChatGPT lazy-loads messages based on scroll position.
 * Solution: Rely on scroll-triggered restoration hooks instead of expand clicks.
 * 
 * This function is a safe no-op that logs diagnostics without side effects.
 */
async function expandCollapsedChatGPTMessages() {
  try {
    console.log('[ExpandChatGPT] Checking for truly collapsed message content...');
    
    // ChatGPT's actual message structure: messages are never "collapsed" in the traditional sense
    // They're either: (1) in viewport and visible, (2) out of viewport and will be lazy-loaded on scroll
    // The aria-expanded="false" elements on the page are UI controls, NOT message containers
    
    // Safe diagnostic only - no clicks
    const messageCount = document.querySelectorAll('[role="article"]').length;
    const visibleText = document.body.innerText.trim().length;
    
    console.log(`[ExpandChatGPT] Found ${messageCount} message container(s), ${visibleText} chars visible`);
    console.log('[ExpandChatGPT] ✓ No unsafe expansions needed - will restore with scroll hooks');
    
    return 0; // No expansions performed
  } catch (e) {
    console.warn('[ExpandChatGPT] Diagnostic error:', e);
    return 0;
  }
}

/* ==================== RESTORE TAGS ==================== */

function restoreTags() {
  if (window.__holorunRestoring) {
    console.log("[Restore] Already restoring, skipping...");
    return;
  }
  
  // TRY to restore even if ping failed - storage API might still work
  // Only skip if we get actual access errors during storage operations
  
  window.__holorunRestoring = true;
  // Start cooldown from the moment a restore actually begins
  try { lastRestoreTime = Date.now(); } catch (_) {}
  
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("🔄 RESTORING TAG HIGHLIGHTS");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");
  
  // First, expand any collapsed content sections to make content accessible
  expandCollapsedChatGPTMessages().then(() => {
    expandCollapsedSections(); // Also expand content sections on any site
    try {
      safeStorageGet({ tags: [] }, (res) => {
      const tags = res.tags || [];
      const currentUrl = cleanUrl(location.href);
      
      console.log("📦 Total tags in storage:", tags.length);
      console.log("📍 Current URL (cleaned):", currentUrl);
      
      const pageTags = tags.filter(t => samePageUrl(t.url, currentUrl));
      console.log("🎯 Tags for this page:", pageTags.length);
      if (pageTags.length > 0 && !samePageUrl(pageTags[0].url, currentUrl)) {
        console.log("   ✅ Tags matched using origin+pathname matching");
      }
      
      let cachedCount = 0;
      let restoredCount = 0;
      
      // Guard highlight creation with mutation observer flag
      window.__holorunCreatingHighlight = true;
      
      pageTags.forEach(tag => {
        // Skip if highlight already exists in DOM (check multiple times to catch highlights created during loop)
        const existing = document.querySelector(`[data-tag-id="${tag.id}"]`);
        if (existing) {
          cachedCount++;
          restoredCount++;
          if (typeof cacheHighlightPosition === 'function') {
            cacheHighlightPosition(tag);
          }
          return;
        }
        
        // Try to find and restore highlight
        let range = null;
        
        console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.log("🔄 RESTORING TAG");
        console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.log(`📌 Tag ID: ${tag.id}`);
        console.log(`🏷️  Topic: "${tag.topic}"`);
        console.log(`📝 Text Type: ${tag.kind || 'unknown'}`);
        console.log(`📏 Text Length: ${tag.text.length} chars`);
        console.log(`📍 Container: <${tag.containerTag || 'unknown'}>`);
        console.log(`📄 Text Preview: "${tag.text.substring(0, 100)}${tag.text.length > 100 ? '...' : ''}"`); 
        
        // Icon tags
        if (tag.iconMeta) {
          console.log(`🎨 Icon Tag - Searching for: label="${tag.iconMeta.label}", ariaLabel="${tag.iconMeta.ariaLabel}"`);
          const iconEl = findIconElement(tag.iconMeta);
          if (iconEl) {
            console.log(`✅ Icon element found at:`, iconEl);
            const highlight = wrapElementWithHighlight(iconEl, tag);
            if (highlight) {
              console.log(`✅ Icon highlight created successfully`);
              console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
              cachedCount++;
              restoredCount++;
              if (typeof cacheHighlightPosition === 'function') {
                cacheHighlightPosition(tag);
              }
              return;
            }
          } else {
            console.log(`❌ Icon element not found`);
          }
        }
        
        // Text tags: route to kind-specific matcher first, then fallback strategies
        const matchText = getMatchableText(tag);
        const kind = tag.kind || classifyTagKind(matchText, tag.containerTag || '', tag.iconMeta);
        console.log(`🔍 Attempting to match text (${matchText.length} chars)`);
        
        // Kind-specific primary strategies
        switch (kind) {
          case 'bullet':
          case 'bullet+icon':
            console.log(`📍 Primary: LI-scoped bullet search`);
            range = findRangeInListItem(matchText);
            break;
          case 'header':
          case 'header+number':
          case 'header+icon':
          case 'header+number+icon':
            console.log(`📍 Primary: Header element search`);
            range = findRangeInHeader(matchText);
            break;
          case 'text+numbers':
            console.log(`📍 Primary: Number-flexible search`);
            range = findRangeByRawTextNumbers(matchText);
            break;
          default:
            console.log(`📍 Primary: Context-based search (anchors: before=${tag.context?.before ? 'yes' : 'no'}, after=${tag.context?.after ? 'yes' : 'no'})`);
            range = findRangeByRawText(matchText, tag.context);
        }
        
        // Fallback strategies if kind-specific matcher failed
        if (!range) {
          console.log(`📍 Fallback 1: Context-based search`);
          range = findRangeByRawText(matchText, tag.context);
        }
        if (!range) {
          console.log(`📍 Fallback 2: Token-based cross-node search`);
          range = findRangeByTokens(matchText, tag.id);
        }
        if (!range) {
          console.log(`📍 Fallback 3: Loose text search (no anchors)`);
          range = findRangeByRawTextLoose(matchText);
        }
        
        if (range) {
          // Log range details BEFORE trying to highlight
          const rangeText = range.toString();
          console.log(`✅ Match found!`);
          console.log(`📏 Range Length: ${rangeText.length} chars`);
          console.log(`📄 Range Text: "${rangeText.substring(0, 100)}${rangeText.length > 100 ? '...' : ''}"`);
          console.log(`📍 Range Location:`, {
            startContainer: range.startContainer.nodeName,
            startOffset: range.startOffset,
            endContainer: range.endContainer.nodeName,
            endOffset: range.endOffset,
            commonAncestor: range.commonAncestorContainer.nodeName
          });
          
          const highlight = surroundWithHighlight(range, tag);
          if (highlight) {
            console.log(`✅ Highlight applied successfully`);
            console.log(`🎨 Highlight Element:`, highlight);
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
            restoredCount++;
          } else {
            console.log(`❌ Highlight creation failed (range may have been rejected by validation)`);
            console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
          }
          cachedCount++;
          if (typeof cacheHighlightPosition === 'function') {
            cacheHighlightPosition(tag);
          }
        } else {
          console.log(`❌ No match found after trying all strategies`);
          console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
        }
      });

      setTimeout(() => {
        console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━");
        console.log("✓ Restore complete!");
        console.log("Found & highlighted:", restoredCount, "/", pageTags.length);
        console.log("Cached:", cachedCount, "/", pageTags.length);
        console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
        renderPanel();
        
        // Handle deep-linking: jump to highlight if #holorunTagId=<id> in URL
        handleDeepLink();
        
        setTimeout(() => {
          window.__holorunCreatingHighlight = false;
          window.__holorunRestoring = false;
          console.log("[Restore] Highlights restored and ready");
        }, 300);
      }, 500);
      });
    } catch (e) {
      console.warn('[RestoreTags] Extension context lost:', e.message);
      window.__holorunRestoring = false;
    }
  }).catch(e => {
    console.warn('[RestoreTags] Expand operation failed:', e.message);
    window.__holorunRestoring = false;
  });
}

/* ==================== STORAGE SYNC ==================== */

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes.tags) {
    console.log("Storage changed");
    renderPanel();
    restoreTags();
  }
});

/* ==================== CONTEXT MIGRATION ==================== */

function migrateAddContext() {
  return new Promise((resolve) => {
    const MIGRATION_FLAG = 'holorunContextMigrated';
    // If extension context is invalid, skip migration and mark flag in localStorage to avoid repeats
    if (!isExtensionContextValid()) {
      try { localStorage.setItem(MIGRATION_FLAG, 'true'); } catch (_) {}
      resolve();
      return;
    }
    
    // Check if already migrated
    safeStorageGet({ [MIGRATION_FLAG]: false }, (migrated) => {
      if (migrated[MIGRATION_FLAG]) {
        resolve();
        return;
      }

      safeStorageGet({ tags: [] }, (result) => {
        const tags = result.tags || [];
        let migrationCount = 0;

        const migratedTags = tags.map(tag => {
          // If tag already has context, skip it
          if (tag.context && (tag.context.before || tag.context.after)) {
            return tag;
          }
          
          // Add empty context for backward compatibility
          migrationCount++;
          return {
            ...tag,
            context: { before: '', after: '' }
          };
        });

        if (migrationCount > 0) {
          safeStorageSet({ tags: migratedTags }, () => {
            console.log(`✅ Added context to ${migrationCount} tag(s)`);
            safeStorageSet({ [MIGRATION_FLAG]: true }, () => resolve());
          });
        } else {
          safeStorageSet({ [MIGRATION_FLAG]: true }, () => resolve());
        }
      });
    });
  });
}

/* ==================== INIT ==================== */

console.log("Initializing...");

// Wait for ChatGPT content to actually load
function waitForChatGPTContent() {
  return new Promise((resolve) => {
    const startTime = Date.now();
    const host = location.hostname || '';
    const isChatGPT = /chatgpt\.com$|chat\.openai\.com$/.test(host);
    const isClaude = /claude\.ai$/.test(host);
    console.log(`[ContentCheck] Site: ${host} - isChatGPT: ${isChatGPT}, isClaude: ${isClaude}`);
    const MAX_WAIT = (isChatGPT || isClaude) ? 15000 : 8000; // Wait longer for chat apps
    const MIN_CONTENT = isChatGPT ? 15000 : isClaude ? 6000 : 2000; // Claude needs more than static pages
    let scrollAttempted = false;
    
    function checkContent() {
      const elapsed = Date.now() - startTime;
      
      // Linearize to see how much content is actually available
      const { full } = linearizeTextNodes(document.body);
      const contentLength = full ? full.length : 0;
      
      console.log(`[ContentCheck] ${contentLength} chars available (${Math.round(elapsed/1000)}s elapsed)`);
      
      // Diagnostic: Log DOM structure on first check to understand ChatGPT layout
      if (elapsed < 1000) {
        const main = document.querySelector('main');
        const roleMain = document.querySelector('[role="main"]');
        
        console.log(`[ContentCheck] 🔍 DOM Diagnostics:`, {
          main: main ? {
            exists: true,
            scrollHeight: main.scrollHeight,
            clientHeight: main.clientHeight,
            isScrollable: main.scrollHeight > main.clientHeight,
            children: main.children.length,
            classes: main.className
          } : { exists: false },
          roleMain: roleMain ? {
            exists: true,
            scrollHeight: roleMain.scrollHeight,
            clientHeight: roleMain.clientHeight,
            isScrollable: roleMain.scrollHeight > roleMain.clientHeight
          } : { exists: false },
          body: {
            scrollHeight: document.body.scrollHeight,
            clientHeight: document.body.clientHeight
          },
          documentElement: {
            scrollHeight: document.documentElement.scrollHeight,
            clientHeight: document.documentElement.clientHeight
          }
        });
      }
      
      // Success: enough content loaded
      if (contentLength >= MIN_CONTENT) {
        console.log(`✓ Content ready: ${contentLength} chars`);
        resolve();
        return;
      }
      
      // If content hasn't loaded after 3 seconds, try progressive scrolling to trigger lazy-load
      if (elapsed >= 3000 && !scrollAttempted && (isChatGPT || isClaude)) {
        scrollAttempted = true;
        console.log(`[ContentCheck] 🔄 Starting progressive scroll to force content load...`);
        console.log(`[ContentCheck] Site check: ${host} - isChatGPT: ${isChatGPT}, isClaude: ${isClaude}`);
        
        // Find scrollable container (conversation area)
        const scrollableSelectors = [
          'main',
          '[role="main"]',
          'main > div',
          'main > div > div'
        ];
        
        let scrollContainer = null;
        for (const selector of scrollableSelectors) {
          const elem = document.querySelector(selector);
          if (elem && elem.scrollHeight > elem.clientHeight) {
            scrollContainer = elem;
            console.log(`[ContentCheck] ✓ Found scrollable container: ${selector}`, {
              scrollHeight: elem.scrollHeight,
              clientHeight: elem.clientHeight,
              canScroll: elem.scrollHeight - elem.clientHeight
            });
            break;
          }
        }
        
        if (!scrollContainer) {
          console.log(`[ContentCheck] ⚠️ No scrollable container found, using window scroll`);
          scrollContainer = window;
        }
        
        // Progressive scroll: simulate user scrolling from top to bottom
        (async () => {
          const scrollStep = 500; // Scroll 500px at a time
          const scrollDelay = 150; // Wait 150ms between scrolls
          const maxScrollHeight = scrollContainer === window 
            ? document.documentElement.scrollHeight 
            : scrollContainer.scrollHeight;
          
          let currentScroll = 0;
          let scrollIterations = 0;
          
          console.log(`[ContentCheck] 📜 Progressive scroll starting (total height: ${maxScrollHeight}px, will scroll ${scrollStep}px per step)`);
          
          // Scroll down in increments
          while (currentScroll < maxScrollHeight && scrollIterations < 50) {
            currentScroll += scrollStep;
            scrollIterations++;
            
            if (scrollContainer === window) {
              window.scrollTo(0, currentScroll);
            } else {
              scrollContainer.scrollTop = currentScroll;
            }
            
            // Log every 5 iterations
            if (scrollIterations % 5 === 0) {
              console.log(`[ContentCheck] 📍 Scrolled to ${currentScroll}px (iteration ${scrollIterations})`);
            }
            
            // Wait for content to load
            await new Promise(resolve => setTimeout(resolve, scrollDelay));
            
            // Check if we've loaded enough content
            const { full } = linearizeTextNodes(document.body);
            const chars = full ? full.length : 0;
            
            if (chars >= MIN_CONTENT) {
              console.log(`[ContentCheck] ✅ Content loaded during scroll at ${currentScroll}px: ${chars} chars`);
              // Scroll back to top
              if (scrollContainer === window) {
                window.scrollTo(0, 0);
              } else {
                scrollContainer.scrollTop = 0;
              }
              console.log(`[ContentCheck] ↩️ Scrolled back to top`);
              resolve();
              return;
            }
          }

          // If chat app, also scroll UP to load older messages
          if (isChatGPT || isClaude) {
            console.log(`[ContentCheck] 🔄 Scrolling upward to load older messages...`);
            while (currentScroll > 0 && scrollIterations < 100) {
              currentScroll -= scrollStep;
              if (currentScroll < 0) currentScroll = 0;
              scrollIterations++;

              if (scrollContainer === window) {
                window.scrollTo(0, currentScroll);
              } else {
                scrollContainer.scrollTop = currentScroll;
              }

              if (scrollIterations % 5 === 0) {
                console.log(`[ContentCheck] 📍 Scrolled up to ${currentScroll}px (iteration ${scrollIterations})`);
              }

              await new Promise(resolve => setTimeout(resolve, scrollDelay));

              const { full } = linearizeTextNodes(document.body);
              const chars = full ? full.length : 0;
              if (chars >= MIN_CONTENT) {
                console.log(`[ContentCheck] ✅ Content loaded during upward scroll at ${currentScroll}px: ${chars} chars`);
                if (scrollContainer === window) {
                  window.scrollTo(0, 0);
                } else {
                  scrollContainer.scrollTop = 0;
                }
                console.log(`[ContentCheck] ↩️ Scrolled back to top`);
                resolve();
                return;
              }
            }
          }

          console.log(`[ContentCheck] 🏁 Finished scrolling (${scrollIterations} iterations), scrolling back to top...`);
          if (scrollContainer === window) {
            window.scrollTo(0, 0);
          } else {
            scrollContainer.scrollTop = 0;
          }
        })();
      }
      
      // Timeout: proceed anyway to avoid infinite wait
      if (elapsed >= MAX_WAIT) {
        console.log(`⚠️ Timeout waiting for content (${contentLength} chars after ${MAX_WAIT}ms)`);
        console.log(`⚠️ ChatGPT may have collapsed conversation. Restoration will retry on scroll.`);
        resolve();
        return;
      }
      
      // Keep waiting
      setTimeout(checkContent, 500);
    }
    
    checkContent();
  });
}

function init() {
  console.log("🚀 Running initialization...");
  console.log("URL:", location.href);
  console.log("Document ready state:", document.readyState);

  // Render panel UI immediately so the FAB shows before content loads
  if (typeof renderPanel === 'function') {
    renderPanel();
  }
  
  // Expand collapsed CONTENT sections (not UI controls) early so content is available for restoration
  expandCollapsedSections();
  
  setTimeout(() => {
    console.log("⏰ Waiting for ChatGPT content to load...");
    waitForChatGPTContent().then(() => {
      console.log("🚀 Starting tag restoration");
      Promise.all([
        migrateLegacyTags(), // One-time migration for legacy tags
        migrateAddContext()  // One-time migration for context field
      ]).then(() => {
        lastRestoreTime = Date.now(); // Set timestamp before first restore
        restoreTags();
        setTimeout(autoJumpIfTagIdPresent, 1200);
      }).catch(err => {
        console.warn("Migration error:", err);
        lastRestoreTime = Date.now();
        restoreTags();
        setTimeout(autoJumpIfTagIdPresent, 1200);
      });
    });
  }, 1000); // Initial 1s delay before starting content check
}

if (document.readyState === "loading") {
  console.log("⏳ Waiting for DOMContentLoaded...");
  document.addEventListener("DOMContentLoaded", init);
} else if (document.readyState === "interactive" || document.readyState === "complete") {
  console.log("✓ DOM already loaded");
  init();
}

console.log("✓ Script loaded successfully");

/* ==================== DEBUG HELPERS ==================== */

function toggleTagPanel() {
  const panelHost = document.getElementById("tag-panel-host");
  if (!panelHost) {
    console.log("[toggleTagPanel] Panel not created yet");
    return;
  }
  const isHidden = panelHost.style.display === 'none';
  panelHost.style.display = isHidden ? 'fixed !important' : 'none !important';
  console.log(`[toggleTagPanel] Panel ${isHidden ? 'shown' : 'hidden'}`);
}

window.holorunTagger = {
  togglePanel: function() {
    toggleTagPanel();
  },

  clearAllTags: function() {
    if (confirm("Delete ALL tags and highlights? This cannot be undone!")) {
      // Remove all highlight spans from the DOM
      document.querySelectorAll('.tag-highlight').forEach(e => e.remove());
      // Remove all pending highlights
      document.querySelectorAll('.pending-highlight').forEach(e => e.remove());
      // Clear tags from storage
      safeStorageSet({ tags: [] }, () => {
        console.log("✅ All tags deleted!");
        alert("All tags and highlights cleared!");
        // Optionally, re-render the panel
        if (typeof renderPanel === 'function') renderPanel();
      });
    }
  },
  
  viewAllTags: function() {
    safeStorageGet({ tags: [] }, (res) => {
      console.log("📦 Total tags:", res.tags.length);
      res.tags.forEach((tag, idx) => {
        console.log(`\n[${idx + 1}] ${tag.topic}`);
        console.log(`  Text: "${tag.text.substring(0, 50)}..."`);
        console.log(`  ID: ${tag.id}`);
        console.log(`  Has context: ${tag.context && (tag.context.before || tag.context.after) ? '✓ YES' : '✗ NO'}`);
        if (tag.context && (tag.context.before || tag.context.after)) {
          console.log(`    Before: "${tag.context.before}"`);
          console.log(`    After: "${tag.context.after}"`);
        }
      });
      console.table(res.tags);
    });
  },
  
  deleteTagsByURL: function(url) {
    safeStorageGet({ tags: [] }, (res) => {
      const filtered = res.tags.filter(t => t.url !== url);
      safeStorageSet({ tags: filtered }, () => {
        console.log("✅ Deleted tags from:", url);
        location.reload();
      });
    });
  },
  
  restoreNow: function() {
    restoreTags();
  },
  
  jumpTo: function(tagId) {
    jumpToHighlight(tagId);
  },

  diagnoseHighlights: function() {
    console.log("\n╔════════════════════════════════════════╗");
    console.log("║  HIGHLIGHT DIAGNOSTIC REPORT           ║");
    console.log("╚════════════════════════════════════════╝");
    
    // Find all highlights in DOM
    const highlights = document.querySelectorAll('.tag-highlight');
    console.log(`\n📊 Total highlights in DOM: ${highlights.length}`);
    
    if (highlights.length === 0) {
      console.log("⚠️  No highlights found in DOM");
      return;
    }
    
    highlights.forEach((span, idx) => {
      console.log(`\n[Highlight ${idx + 1}]`);
      console.log(`  Tag ID: ${span.dataset.tagId}`);
      console.log(`  Topic: ${span.dataset.tag}`);
      console.log(`  Text: "${span.textContent.substring(0, 50)}..."`);
      
      const styles = window.getComputedStyle(span);
      console.log(`  CSS Applied:`);
      console.log(`    - Background: ${styles.backgroundColor}`);
      console.log(`    - Display: ${styles.display}`);
      console.log(`    - Visibility: ${styles.visibility}`);
      console.log(`    - Opacity: ${styles.opacity}`);
      console.log(`    - Padding: ${styles.padding}`);
      console.log(`    - Border-Radius: ${styles.borderRadius}`);
      
      console.log(`  Visibility Check:`);
      console.log(`    - In DOM: ✓ YES`);
      console.log(`    - Visible: ${styles.visibility !== 'hidden' && styles.opacity !== '0' ? '✓ YES' : '✗ NO'}`);
      console.log(`    - Parent visible: ${span.parentElement && window.getComputedStyle(span.parentElement).visibility !== 'hidden' ? '✓ YES' : '✗ NO'}`);
      
      // Check if clickable
      console.log(`  Interactivity:`);
      console.log(`    - Has click handler: ${span.onclick ? '✓ YES' : '✗ NO'}`);
      console.log(`    - Pointer events: ${styles.pointerEvents}`);
    });
    
    // Summary
    console.log(`\n📈 Summary:`);
    console.log(`  Total highlights: ${highlights.length}`);
    console.log(`  Visible highlights: ${Array.from(highlights).filter(h => {
      const styles = window.getComputedStyle(h);
      return styles.visibility !== 'hidden' && styles.opacity !== '0';
    }).length}`);
  },

  testRangeMatching: function(targetText = null) {
    console.log("\n╔════════════════════════════════════════╗");
    console.log("║  RANGE MATCHING DIAGNOSTIC             ║");
    console.log("╚════════════════════════════════════════╝");
    
    if (!targetText) {
      console.log("\n📝 Usage: window.holorunTagger.testRangeMatching('text to find')");
      console.log("Examples:");
      safeStorageGet({ tags: [] }, (res) => {
        if (res.tags.length > 0) {
          console.log("\nAvailable tags:");
          res.tags.slice(0, 3).forEach((tag, i) => {
            console.log(`  ${i + 1}. "${tag.text.substring(0, 40)}..."`);
            console.log(`     window.holorunTagger.testRangeMatching('${tag.text}')`);
          });
        } else {
          console.log("No tags in storage. Create a tag first.");
        }
      });
      return;
    }

    console.log(`\nSearching for: "${targetText.substring(0, 60)}..."`);
    
    // Test linearization
    console.log(`\n1️⃣ Linearization Test:`);
    const { full, mapping } = linearizeTextNodes(document.body);
    console.log(`  Total chars linearized: ${full.length}`);
    console.log(`  Total text nodes found: ${mapping.length}`);
    
    if (full.length > 0 && targetText) {
      const idx = full.toLowerCase().indexOf(targetText.toLowerCase());
      if (idx >= 0) {
        console.log(`  ✅ Text found at position ${idx}`);
        console.log(`  Preview: "...${full.substring(Math.max(0, idx - 20), idx + 50)}..."`);
      } else {
        console.log(`  ❌ Text not found in linearized content`);
      }
    }
    
    // Test each matching strategy
    console.log(`\n2️⃣ Testing findRangeByRawText:`);
    const range1 = findRangeByRawText(targetText);
    if (range1) {
      const text1 = range1.toString();
      console.log(`  ✅ Found - length: ${text1.length} chars`);
      console.log(`  Preview: "${text1.substring(0, 60)}..."`);
    } else {
      console.log(`  ❌ Not found`);
    }
    
    console.log(`\n3️⃣ Testing findRangeByTokens:`);
    const range2 = findRangeByTokens(targetText);
    if (range2) {
      const text2 = range2.toString();
      console.log(`  ✅ Found - length: ${text2.length} chars`);
      console.log(`  Preview: "${text2.substring(0, 60)}..."`);
    } else {
      console.log(`  ❌ Not found`);
    }
    
    console.log(`\n4️⃣ Testing findRangeByRawTextLoose:`);
    const range3 = findRangeByRawTextLoose(targetText);
    if (range3) {
      const text3 = range3.toString();
      console.log(`  ✅ Found - length: ${text3.length} chars`);
      console.log(`  Preview: "${text3.substring(0, 60)}..."`);
    } else {
      console.log(`  ❌ Not found`);
    }
    
    console.log(`\n5️⃣ Summary:`);
    const found = [range1, range2, range3].filter(r => r).length;
    console.log(`  Strategies succeeded: ${found}/3`);
    console.log(`\n💡 Tip: Check browser console for detailed logs like [FindRawText], [FindRangeTokens], [FindRangeLoose]`);
  }
};

// --- Diagnostic: Check panel visibility and highlight state ---
window.holorunTagger.diagnosePanel = function() {
  console.log('\n========== PANEL DIAGNOSIS ==========');
  
  // Check Shadow DOM panel host
  const panelHost = document.getElementById('tag-panel-host');
  console.log('1️⃣ Panel Host:', panelHost ? '✅ EXISTS' : '❌ NOT FOUND');
  
  if (panelHost) {
    const computed = window.getComputedStyle(panelHost);
    console.log('   - Display:', computed.display);
    console.log('   - Position:', computed.position);
    console.log('   - Right:', computed.right);
    console.log('   - Top:', computed.top);
    console.log('   - Width:', computed.width);
    console.log('   - z-index:', computed.zIndex);
    console.log('   - Transform:', computed.transform);
    
    // Check Shadow DOM content
    const shadowRoot = panelHost.shadowRoot;
    console.log('2️⃣ Shadow Root:', shadowRoot ? '✅ EXISTS' : '❌ NOT FOUND');
    
    if (shadowRoot) {
      const panel = shadowRoot.querySelector('#tag-panel');
      console.log('3️⃣ Panel Container:', panel ? '✅ EXISTS' : '❌ NOT FOUND');
      
      const tagListPane = shadowRoot.querySelector('#tag-list-pane');
      const previewPane = shadowRoot.querySelector('#preview-pane');
      console.log('4️⃣ Tag List Pane:', tagListPane ? '✅ EXISTS' : '❌ NOT FOUND');
      console.log('5️⃣ Preview Pane:', previewPane ? '✅ EXISTS' : '❌ NOT FOUND');
      
      if (tagListPane) {
        const tagRows = tagListPane.querySelectorAll('.tag-row');
        console.log('   - Tag rows found:', tagRows.length);
        if (tagRows.length > 0) {
          console.log('   - First tag text:', tagRows[0].querySelector('.tag-label')?.textContent?.slice(0, 30) + '...');
        }
      }
      
      if (previewPane) {
        const previewComputed = window.getComputedStyle(previewPane);
        console.log('   - Preview display:', previewComputed.display);
        console.log('   - Preview flex:', previewComputed.flex);
        const previewContent = shadowRoot.querySelector('#preview-content');
        console.log('   - Preview content:', previewContent ? '✅ EXISTS' : '❌ NOT FOUND');
        if (previewContent) {
          console.log('   - Content length:', previewContent.innerHTML.length, 'chars');
        }
      }
    }
  }
  
  // Check FAB button
  const fabHost = document.getElementById('fab-host');
  console.log('6️⃣ FAB Host:', fabHost ? '✅ EXISTS' : '❌ NOT FOUND');
  if (fabHost) {
    const fabComputed = window.getComputedStyle(fabHost);
    console.log('   - FAB display:', fabComputed.display);
    console.log('   - FAB position:', fabComputed.position);
    console.log('   - FAB right:', fabComputed.right);
    console.log('   - FAB top:', fabComputed.top);
  }
  
  const highlights = document.querySelectorAll('[data-tag-id]');
  console.log('7️⃣ Highlight elements found:', highlights.length);
  highlights.forEach((h, i) => {
    if (i < 3) { // Only show first 3 to avoid spam
      const computed = window.getComputedStyle(h);
      const rect = h.getBoundingClientRect();
      const inViewport = rect.top >= 0 && rect.top <= window.innerHeight;
      console.log(`   Highlight ${i+1}: ${h.getAttribute('data-tag-id').substring(0,8)}...`);
      console.log(`     - Background: ${computed.backgroundColor}`);
      console.log(`     - Display: ${computed.display}`);
      console.log(`     - Position: top=${Math.round(rect.top)}, left=${Math.round(rect.left)}`);
      console.log(`     - In viewport: ${inViewport ? '✓ YES' : '✗ NO'}`);
      console.log(`     - Parent: ${h.parentElement ? h.parentElement.tagName : 'none'}`);
    }
  });
  
  console.log('========== END DIAGNOSIS ==========\n');
};

// --- Diagnostic: Test preview pane functionality ---
window.holorunTagger.diagnosePreview = function() {
  console.log('\n========== PREVIEW PANE DIAGNOSIS ==========');
  
  const panelHost = document.getElementById('tag-panel-host');
  if (!panelHost || !panelHost.shadowRoot) {
    console.log('❌ Panel not found or no shadow root');
    return;
  }
  
  const shadowRoot = panelHost.shadowRoot;
  const tagListPane = shadowRoot.querySelector('#tag-list-pane');
  const previewPane = shadowRoot.querySelector('#preview-pane');
  const previewContent = shadowRoot.querySelector('#preview-content');
  
  console.log('1️⃣ Layout Elements:');
  console.log('   - Tag List Pane:', tagListPane ? '✅ EXISTS' : '❌ NOT FOUND');
  console.log('   - Preview Pane:', previewPane ? '✅ EXISTS' : '❌ NOT FOUND');
  console.log('   - Preview Content:', previewContent ? '✅ EXISTS' : '❌ NOT FOUND');
  
  if (previewPane) {
    const computed = window.getComputedStyle(previewPane);
    console.log('2️⃣ Preview Pane Styles:');
    console.log('   - Display:', computed.display);
    console.log('   - Flex:', computed.flex);
    console.log('   - Width:', computed.width);
    console.log('   - Height:', computed.height);
    console.log('   - Background:', computed.background);
    console.log('   - Border:', computed.border);
  }
  
  if (previewContent) {
    console.log('3️⃣ Preview Content:');
    console.log('   - Content length:', previewContent.innerHTML.length, 'characters');
    console.log('   - Has content:', previewContent.innerHTML.length > 100 ? '✅ YES' : '🔶 MINIMAL/EMPTY');
    if (previewContent.innerHTML.length > 0) {
      const hasButtons = previewContent.querySelectorAll('button').length;
      console.log('   - Action buttons:', hasButtons, 'found');
    }
  }
  
  if (tagListPane) {
    const rows = tagListPane.querySelectorAll('.tag-row');
    console.log('4️⃣ Tag List:');
    console.log('   - Rows found:', rows.length);
    if (rows.length > 0) {
      const hasHoverHandler = rows[0].onmouseenter !== null;
      const hasClickHandler = rows[0].onclick !== null;
      console.log('   - First row has hover:', hasHoverHandler ? '✅ YES' : '❌ NO');
      console.log('   - First row has click:', hasClickHandler ? '✅ YES' : '❌ NO');
    }
  }
  
  console.log('5️⃣ Test Preview Function:');
  if (typeof window.showTagPreview === 'function') {
    console.log('   - showTagPreview function: ✅ EXISTS');
  } else {
    console.log('   - showTagPreview function: ❌ NOT FOUND');
  }
  
  console.log('\n💡 Try: Hover over a tag to test preview functionality');
};

// --- Diagnostic: Test scroll to a specific highlight ---
window.holorunTagger.testScroll = function(tagId) {
  console.log('\n========== SCROLL TEST ==========');
  
  if (!tagId) {
    console.log('Usage: window.holorunTagger.testScroll("tag-id")');
    return;
  }
  
  const highlight = document.querySelector(`[data-tag-id="${tagId}"]`);
  if (!highlight) {
    console.log('❌ Highlight not found');
    return;
  }
  
  const before = highlight.getBoundingClientRect();
  console.log(`Before scroll: top=${before.top}, left=${before.left}`);
  
  highlight.scrollIntoView({ behavior: 'auto', block: 'center', inline: 'nearest' });
  
  setTimeout(() => {
    const after = highlight.getBoundingClientRect();
    console.log(`After scroll: top=${after.top}, left=${after.left}`);
    console.log(`Scrolled: ${before.top !== after.top ? '✓ YES' : '✗ NO'}`);
    console.log(`In viewport: ${after.top >= 0 && after.top <= window.innerHeight ? '✓ YES' : '✗ NO'}`);
    console.log('========== END SCROLL TEST ==========\n');
  }, 500);
};

console.log("💡 Debug commands available:");
console.log("  window.holorunTagger.clearAllTags() - Delete all tags");
console.log("  window.holorunTagger.viewAllTags() - View all tags");
console.log("  window.holorunTagger.restoreNow() - Try restoring tags now");
console.log("  window.holorunTagger.jumpTo('tag-id') - Jump to specific tag");
console.log("  window.holorunTagger.diagnoseHighlights() - Diagnose all highlights & CSS");
console.log("  window.holorunTagger.diagnosePanel() - Diagnose panel visibility & colors");
console.log("  window.holorunTagger.diagnosePreview() - Diagnose preview pane layout");
console.log("  window.holorunTagger.testScroll('tag-id') - Test scroll to a specific highlight");
console.log("  window.holorunTagger.testRangeMatching('text') - Test text matching strategies");
console.log("  window.holorunTagger.diagnoseWorkflow() - Diagnose workflow grid rendering");
