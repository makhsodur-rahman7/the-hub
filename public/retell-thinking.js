(function() {
    let thinkingTimeout = null;
    let observer = null;
    let messagesObserver = null;
    let isThinking = false;
    let currentInput = null;
    let currentButton = null;
    
    let originalPlaceholder = ''; 
    
    function setThinkingState(input, button, isThinkingNow) {
        if (!input) return;
        isThinking = isThinkingNow;
        if (isThinkingNow) {
            if (input.placeholder && input.placeholder !== 'Thanks, bear with me...') {
                originalPlaceholder = input.placeholder;
            }
            input.placeholder = 'Thanks, bear with me...';
            input.disabled = true;
            if (button) button.disabled = true;
            
            if (thinkingTimeout) clearTimeout(thinkingTimeout);
            // Safety fallback: 30 seconds
            thinkingTimeout = setTimeout(() => {
                if (isThinking) setThinkingState(input, button, false);
            }, 30000);
        } else {
            input.placeholder = originalPlaceholder || 'Type your message...';
            input.disabled = false;
            if (button) button.disabled = false;
            if (thinkingTimeout) clearTimeout(thinkingTimeout);
        }
    }

    function initAddon(shadowRoot) {
        // Fallback selectors if Retell updates their exact classes
        const input = shadowRoot.querySelector('textarea') || shadowRoot.querySelector('input[type="text"][class*="_input_"]');
        const button = shadowRoot.querySelector('button[class*="_sendButton_"]') || shadowRoot.querySelector('button[type="submit"]');
        const messagesContainer = shadowRoot.querySelector('[class*="_messages_"]');
        
        if (!input || !messagesContainer) {
            // console.log('[Retell Thinking] Missing elements:', { input: !!input, messagesContainer: !!messagesContainer });
            return;
        }
        
        // If we already attached to this specific input, do nothing
        if (currentInput === input) return;
        
        console.log('[Retell Thinking] Successfully found elements and attached listeners');
        currentInput = input;
        currentButton = button;
        
        // 1. Listen for Enter key
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                if (input.value.trim() !== '') {
                    console.log('[Retell Thinking] Triggered by Enter');
                    // Small delay so Retell's own React handler fires first before we disable the input
                    setTimeout(() => setThinkingState(input, button, true), 50);
                }
            }
        });
        
        // 2. Listen for Send button click
        if (button) {
            button.addEventListener('click', () => {
                if (input.value.trim() !== '') {
                    console.log('[Retell Thinking] Triggered by Button click');
                    setTimeout(() => setThinkingState(input, button, true), 50);
                }
            });
        }
        
        // Observe messages container for new agent messages
        if (messagesObserver) messagesObserver.disconnect();
        messagesObserver = new MutationObserver((mutations) => {
            for (let mut of mutations) {
                if (mut.type === 'childList') {
                    for (let node of mut.addedNodes) {
                        if (node.nodeType === 1) { // Element node
                            const isAgent = node.matches('[class*="_agent_"]') || node.querySelector('[class*="_agent_"]');
                            if (isAgent && isThinking) {
                                console.log('[Retell Thinking] Found agent reply, restoring state');
                                setThinkingState(currentInput, currentButton, false);
                            }
                        }
                    }
                }
            }
        });
        messagesObserver.observe(messagesContainer, { childList: true, subtree: true });
    }
    
    function setupObservers() {
        console.log('[Retell Thinking] Setting up observers');
        // The widget might be inside a retell-widget tag
        const getWidget = () => Array.from(document.querySelectorAll('*')).find(el => el.shadowRoot && el.tagName.toLowerCase().includes('retell'));
        
        const widget = getWidget();
        if (widget && widget.shadowRoot) {
            console.log('[Retell Thinking] Found widget immediately', widget);
            // Initialize immediately if chat is already open
            initAddon(widget.shadowRoot);
            
            // Keep observing because chat window might open/close, recreating the DOM inside shadowRoot
            observer = new MutationObserver(() => {
                initAddon(widget.shadowRoot);
            });
            observer.observe(widget.shadowRoot, { childList: true, subtree: true });
        } else {
            console.log('[Retell Thinking] Widget not found yet, observing body');
            // Widget tag not found yet, observe document body
            const bodyObserver = new MutationObserver(() => {
                const w = getWidget();
                if (w && w.shadowRoot) {
                    console.log('[Retell Thinking] Found widget after mutation', w);
                    bodyObserver.disconnect();
                    initAddon(w.shadowRoot);
                    observer = new MutationObserver(() => {
                        initAddon(w.shadowRoot);
                    });
                    observer.observe(w.shadowRoot, { childList: true, subtree: true });
                }
            });
            bodyObserver.observe(document.body, { childList: true, subtree: true });
        }
    }
    
    // Start on DOMContentLoaded or immediately if already loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', setupObservers);
    } else {
        setupObservers();
    }
})();
