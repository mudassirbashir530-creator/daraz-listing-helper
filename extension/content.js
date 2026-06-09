/**
 * LIFEINNO Daraz Helper - Content Script
 * Automatically binds to seller.daraz.pk pages for active input fill integrations
 */

let activeFloatingBtn = null;
let currentFocusedElement = null;

// Listen for focus inside input elements on Daraz Seller portal
document.addEventListener('focusin', (e) => {
  const target = e.target;
  if (!target) return;

  // Check if target is a text/number input or textarea
  const isInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA';
  const isReadOnly = target.readOnly || target.disabled;
  const isExcludedType = ['submit', 'button', 'checkbox', 'radio', 'file', 'hidden'].includes(target.type);

  if (isInput && !isReadOnly && !isExcludedType) {
    currentFocusedElement = target;
    showFillFloatingButton(target);
  }
});

// Remove floating button when clicking elsewhere, with a small timeout to let button click trigger first
document.addEventListener('focusout', (e) => {
  setTimeout(() => {
    // If focus didn't move to the floating button, remove it
    if (document.activeElement !== activeFloatingBtn) {
      removeFloatingButton();
    }
  }, 180);
});

// Handle incoming messages from the popup or sidebar panels
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'FILL_FIELD') {
    const targetEl = currentFocusedElement || document.activeElement;
    if (targetEl && (targetEl.tagName === 'INPUT' || targetEl.tagName === 'TEXTAREA')) {
      fillValue(targetEl, msg.value);
      sendResponse({ success: true });
    } else {
      sendResponse({ success: false, error: "No active text field focused" });
    }
  }
  return true; // Keep channel open for async response
});

// Create and position the elegant floating indicator
function showFillFloatingButton(element) {
  removeFloatingButton(); // Clean up existing

  // Create Button
  const btn = document.createElement('button');
  btn.id = 'lifeinno-floating-fill-btn';
  btn.innerHTML = '⚡ Fill';
  btn.title = 'Fill with last copied value from LIFEINNO Daraz Helper';
  
  // Custom Styles for floating button
  Object.assign(btn.style, {
    position: 'absolute',
    background: '#E85D24', // LIFEINNO Orange
    color: '#FFF',
    border: 'none',
    borderRadius: '4px',
    padding: '4px 8px',
    fontSize: '11px',
    fontWeight: 'bold',
    cursor: 'pointer',
    zIndex: '1000000',
    boxShadow: '0 2px 5px rgba(0,0,0,0.2)',
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif",
    transition: 'all 0.1s ease',
    pointerEvents: 'auto'
  });

  // Hover feedback
  btn.addEventListener('mouseenter', () => {
    btn.style.background = '#d14f1b';
    btn.style.transform = 'scale(1.05)';
  });
  btn.addEventListener('mouseleave', () => {
    btn.style.background = '#E85D24';
    btn.style.transform = 'scale(1)';
  });

  // Handle Injecting value on Click
  btn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();

    // Query Chrome local storage to get the last copied field item
    chrome.storage.local.get('lastCopiedValue', (data) => {
      if (data.lastCopiedValue) {
        fillValue(element, data.lastCopiedValue);
        
        // Show success splash on floating button
        btn.innerHTML = '✅ Done!';
        btn.style.background = '#10B981';
        setTimeout(() => {
          removeFloatingButton();
        }, 800);
      } else {
        btn.innerHTML = '⚠️ No Copied Data';
        btn.style.background = '#EF4444';
        setTimeout(() => {
          btn.innerHTML = '⚡ Fill';
          btn.style.background = '#E85D24';
        }, 1500);
      }
    });

    // Return focus to input so the user can continue typing
    element.focus();
  });

  document.body.appendChild(btn);
  activeFloatingBtn = btn;

  // Calculate positioning near focused elements
  repositionButton(element, btn);
}

// Positioning logistics to prevent layout breaks
function repositionButton(element, btn) {
  const rect = element.getBoundingClientRect();
  const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
  const scrollTop = window.pageYOffset || document.documentElement.scrollTop;

  // Place inside or on the right side of the input field
  btn.style.left = `${rect.right + scrollLeft - btn.offsetWidth - 4}px`;
  btn.style.top = `${rect.top + scrollTop + (rect.height - btn.offsetHeight) / 2}px`;
  
  // Guard if input too narrow
  if (rect.width < 120) {
    btn.style.left = `${rect.right + scrollLeft + 4}px`;
  }
}

// Clean target node
function removeFloatingButton() {
  if (activeFloatingBtn && activeFloatingBtn.parentNode) {
    activeFloatingBtn.parentNode.removeChild(activeFloatingBtn);
  }
  activeFloatingBtn = null;
}

// Safely fills the input and triggers framework level react state bindings
function fillValue(element, value) {
  if (!element || value === undefined || value === null) return;
  
  // Standard property fill
  element.value = value;
  
  // Fire synthetic events for React/Vue reactive form engines (Daraz uses key reactive stores)
  const inputEvent = new Event('input', { bubbles: true, cancelable: true });
  const changeEvent = new Event('change', { bubbles: true, cancelable: true });
  
  element.dispatchEvent(inputEvent);
  element.dispatchEvent(changeEvent);
  
  // Custom dispatcher for reactive properties if needed
  const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value");
  const nativeTextareaValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value");
  const setter = element.tagName === 'TEXTAREA' ? nativeTextareaValueSetter : nativeInputValueSetter;
  
  if (setter && setter.set) {
    setter.set.call(element, value);
    element.dispatchEvent(inputEvent);
    element.dispatchEvent(changeEvent);
  }
}
