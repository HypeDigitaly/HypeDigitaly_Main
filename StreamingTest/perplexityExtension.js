export const PerplexityReasonerExtension = {
  name: 'PerplexityReasoner',
  type: 'response',
  match: ({ trace }) =>
    trace.type === 'ext_perplexityReasoner' ||
    trace.payload?.name === 'ext_perplexityReasoner',
  render: async ({ trace, element }) => {
    const container = document.createElement('div')
    container.className = 'perplexity-reasoner-container'

    // Add variables to track streaming state
    let isStreaming = false
    let activeReasoningGroup = null

    // Function to format model name
    function formatModelName(model) {
      const modelMap = {
        'sonar-reasoning': 'Sonar Reasoning model',
        'sonar-reasoning-pro': 'Sonar Reasoning Pro model',
      }
      return (
        modelMap[model] ||
        model
          .split('-')
          .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
          .join(' ')
      )
    }

    // Create the base structure first
    container.innerHTML = `
      <style>
        .perplexity-reasoner-container {
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          width: 100%;
          max-width: none;
          margin: 0;
          padding: 0;
          display: flex;
          flex-direction: column;
          gap: 0;
        }
        .reasoning-section {
          background-color: #F9FAFB;
          border-radius: 12px;
          padding: 16px;
          margin: 0;
          width: 100%;
          box-sizing: border-box;
          transition: all 0.3s ease;
        }
        .reasoning-section.collapsed {
          padding: 10px 16px;
          cursor: pointer;
        }
        .reasoning-section.has-answer {
          margin-bottom: 8px;
        }
        .reasoning-section.collapsed .reasoning-content,
        .reasoning-section.collapsed .reasoning-intro {
          display: none;
        }
        .reasoning-header {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-bottom: 12px;
        }
        .reasoning-section.collapsed .reasoning-header {
          margin-bottom: 0;
        }
        .reasoning-icon {
          width: 28px;
          height: 28px;
          display: flex;
          align-items: center;
          justify-content: center;
        }
        .reasoning-icon svg {
          width: 28px;
          height: 28px;
        }
        .reasoning-title-wrapper {
          flex: 1;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }
        .reasoning-title {
          font-size: 13px;
          font-weight: 600;
          color: #111827;
        }
        .reasoning-model {
          font-weight: normal;
          color: #6B7280;
          margin-left: 4px;
        }
        .toggle-icon {
          width: 16px;
          height: 16px;
          color: #6B7280;
          transition: transform 0.3s ease;
        }
        .reasoning-section.collapsed .toggle-icon {
          transform: rotate(-180deg);
        }
        .reasoning-intro {
          display: flex;
          align-items: center;
          gap: 8px;
          font-size: 12px;
          line-height: 1.4;
          color: #6B7280;
          margin-bottom: 12px;
        }
        .loading-dots {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          height: 20px;
        }
        .loading-dots .dot {
          width: 4px;
          height: 4px;
          background-color: #6B7280;
          border-radius: 50%;
          animation: dotPulse 1.5s infinite;
        }
        .loading-dots .dot:nth-child(2) {
          animation-delay: 0.2s;
        }
        .loading-dots .dot:nth-child(3) {
          animation-delay: 0.4s;
        }
        @keyframes dotPulse {
          0%, 100% {
            opacity: 0.4;
            transform: scale(1);
          }
          50% {
            opacity: 1;
            transform: scale(1.3);
          }
        }
        .reasoning-content {
          font-size: 12px;
          line-height: 1.4;
          color: #4B5563;
        }
        .reasoning-step {
          display: flex;
          align-items: flex-start;
          gap: 8px;
          margin-bottom: 8px;
        }
        .step-checkbox {
          width: 14px;
          height: 14px;
          flex-shrink: 0;
          margin-top: 1px;
          position: relative;
        }
        .step-checkbox svg {
          position: absolute;
          top: 0;
          left: 0;
          width: 14px;
          height: 14px;
        }
        .step-content {
          flex: 1;
          font-size: 12px;
          line-height: 1.4;
          padding-top: 1px;
        }
        .step-checkbox .unchecked {
          opacity: 1;
          transition: opacity 0.3s ease;
        }
        .step-checkbox .checked {
          opacity: 0;
          transition: opacity 0.3s ease;
        }
        .step-checkbox.is-checked .unchecked {
          opacity: 0;
        }
        .step-checkbox.is-checked .checked {
          opacity: 1;
        }
        .answer-section {
          padding: 0;
          margin: 0;
          width: 100%;
          box-sizing: border-box;
          opacity: 0;
          height: 0;
          overflow: hidden;
          transition: opacity 0.3s ease;
        }
        .answer-section.visible {
          opacity: 1;
          height: auto;
          overflow: visible;
        }
        .vfrc-message--extension-PerplexityReasoner {
          width: 100% !important;
          max-width: none !important;
        }
        .answer-content {
          font-size: 14px;
          line-height: 1.4;
          margin: 0;
          padding: 0;
        }
      </style>
      <div class="reasoning-section">
        <div class="reasoning-header">
          <div class="reasoning-icon">
            <svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" fill="none" viewBox="0 0 128 128" id="science">
  <ellipse cx="72" cy="48" stroke="#1B1B1B" stroke-width="6" rx="21.076" ry="10.036" transform="rotate(-45 72 48)"></ellipse>
  <ellipse cx="72" cy="48" stroke="#1B1B1B" stroke-width="6" rx="21.076" ry="10.036" transform="rotate(45 72 48)"></ellipse>
  <path fill="#1B1B1B" d="M107 46.7758C107 45.1189 105.657 43.7758 104 43.7758C102.343 43.7758 101 45.1189 101 46.7758H107ZM70.0342 19C71.6911 19 73.0342 17.6569 73.0342 16C73.0342 14.3431 71.6911 13 70.0342 13V19ZM29.8205 72.6634H32.8205C32.8205 71.446 32.0848 70.3493 30.9583 69.8876L29.8205 72.6634ZM89.611 94.2028C90.1916 95.7546 91.9203 96.5419 93.4721 95.9613C95.0238 95.3807 95.8111 93.652 95.2305 92.1002L89.611 94.2028ZM57.1417 91.2742C58.6739 90.6437 59.4048 88.8904 58.7742 87.3583C58.1437 85.8261 56.3904 85.0952 54.8583 85.7258L57.1417 91.2742ZM86.8025 111V114V111ZM23.3165 69.9974L22.1787 72.7733L23.3165 69.9974ZM55.8648 111L55.8648 114L55.8648 111ZM31.6171 53.0102L29.4088 50.9795L31.6171 53.0102ZM33.6399 48.7818L30.6731 48.3368L33.6399 48.7818ZM39.0923 93.2606L38.6153 90.2988L39.0923 93.2606ZM92.4208 93.1515L95.3162 92.3663C95.2918 92.2764 95.2632 92.1876 95.2305 92.1002L92.4208 93.1515ZM101 46.7758C101 61.4182 98.9356 67.2732 96.6185 70.2041L101.325 73.9252C104.992 69.2875 107 61.607 107 46.7758H101ZM24.0975 65.6194L33.8254 55.0409L29.4088 50.9795L19.681 61.5581L24.0975 65.6194ZM30.9583 69.8876L24.4543 67.2216L22.1787 72.7733L28.6826 75.4393L30.9583 69.8876ZM32.8205 85.3624V72.6634H26.8205V85.3624H32.8205ZM47.3879 88.8862L38.6153 90.2988L39.5692 96.2225L48.3417 94.8099L47.3879 88.8862ZM50.8648 103V91.8481H44.8648V103H50.8648ZM86.8025 108L55.8648 108L55.8648 114L86.8025 114V108ZM89.5253 93.9367L91.6282 101.691L97.4191 100.121L95.3162 92.3663L89.5253 93.9367ZM49.0065 94.6223L57.1417 91.2742L54.8583 85.7258L46.723 89.0738L49.0065 94.6223ZM40.6287 32.7137C48.0449 20.3612 59.6888 19 70.0342 19V13C59.2988 13 44.6452 14.3673 35.4846 29.6252L40.6287 32.7137ZM31.7491 41.1633L30.6731 48.3368L36.6067 49.2268L37.6827 42.0534L31.7491 41.1633ZM86.8025 114C94.049 114 99.3157 107.115 97.4191 100.121L91.6282 101.691C92.4903 104.87 90.0964 108 86.8025 108V114ZM19.681 61.5581C16.3813 65.1463 17.6681 70.9244 22.1787 72.7733L24.4543 67.2216C23.81 66.9575 23.6262 66.132 24.0975 65.6194L19.681 61.5581ZM44.8648 103C44.8648 109.075 49.7897 114 55.8648 114L55.8648 108C53.1034 108 50.8648 105.761 50.8648 103H44.8648ZM33.8254 55.0409C35.3122 53.424 36.2809 51.399 36.6067 49.2268L30.6731 48.3368C30.525 49.3242 30.0847 50.2446 29.4088 50.9795L33.8254 55.0409ZM26.8205 85.3624C26.8205 92.1352 32.8825 97.2992 39.5692 96.2225L38.6153 90.2988C35.5759 90.7882 32.8205 88.441 32.8205 85.3624H26.8205ZM35.4846 29.6252C33.2896 33.2812 32.3184 37.368 31.7491 41.1633L37.6827 42.0534C38.207 38.5581 39.0289 35.3782 40.6287 32.7137L35.4846 29.6252ZM96.6185 70.2041C91.4986 76.6802 86.0249 84.6185 89.611 94.2028L95.2305 92.1002C92.9397 85.9776 96.0444 80.6049 101.325 73.9252L96.6185 70.2041Z"></path>
</svg>
          </div>
          <div class="reasoning-title-wrapper">
            <div class="reasoning-title">Reasoning<span class="reasoning-model">with ${formatModelName(
              trace.payload?.model || 'Unknown Model'
            )}</span></div>
            <svg class="toggle-icon" viewBox="0 0 20 20" fill="currentColor">
              <path fill-rule="evenodd" d="M14.707 12.707a1 1 0 01-1.414 0L10 9.414l-3.293 3.293a1 1 0 01-1.414-1.414l4-4a1 1 0 011.414 0l4 4a1 1 0 010 1.414z" clip-rule="evenodd" />
            </svg>
          </div>
        </div>
        <div class="reasoning-intro">
          <span>Thinking</span>
          <div class="loading-dots">
            <div class="dot"></div>
            <div class="dot"></div>
            <div class="dot"></div>
          </div>
        </div>
        <div class="reasoning-content" id="reasoning-content"></div>
      </div>
      <div class="answer-section">
        <div class="answer-content" id="answer-content"></div>
      </div>
    `

    // Add the container to the element first
    element.appendChild(container)

    // Get references to elements
    const reasoningSection = container.querySelector('.reasoning-section')
    const reasoningContent = container.querySelector('#reasoning-content')
    const answerContent = container.querySelector('#answer-content')
    const answerSection = container.querySelector('.answer-section')

    // Check if model name contains "reasoning"
    const modelName = (trace.payload?.model || '').toLowerCase()
    const isReasoningModel = modelName.includes('reasoning')

    // Hide reasoning section for non-reasoning models
    if (!isReasoningModel) {
      reasoningSection.style.display = 'none'
      // Initially hide the container for non-reasoning models
      container.style.display = 'none'
    }

    // Keep track of completed steps globally
    let completedSteps = []

    // Move scrollToBottom outside and improve implementation
    function scrollToBottom() {
      // Use scrollIntoView on our element with padding
      if (element) {
        // Add temporary padding to the bottom
        const originalPadding = element.style.paddingBottom
        element.style.paddingBottom = '80px'

        element.scrollIntoView({
          behavior: 'smooth',
          block: 'end',
          inline: 'nearest',
        })

        // Backup scroll attempt with auto behavior
        setTimeout(() => {
          element.scrollIntoView({
            behavior: 'auto',
            block: 'end',
            inline: 'nearest',
          })

          // Restore original padding after scrolling
          setTimeout(() => {
            element.style.paddingBottom = originalPadding
          }, 300)
        }, 100)
      }
    }

    // Add click handler for reasoning section expansion
    reasoningSection.addEventListener('click', (e) => {
      if (e.target.closest('.reasoning-header')) {
        const isCollapsed = reasoningSection.classList.contains('collapsed')
        if (isCollapsed) {
          // Expanding
          reasoningSection.classList.remove('collapsed')
          // Restore height and visibility
          reasoningContent.style.height = 'auto'
          reasoningContent.style.overflow = 'visible'
          // Show all completed steps and their checkboxes
          completedSteps.forEach((step) => {
            if (step) {
              step.style.display = 'flex'
              const checkbox = step.querySelector('.step-checkbox')
              if (checkbox) {
                checkbox.style.display = 'flex'
                checkbox.classList.add('is-checked')
              }
            }
          })
          // Update height after content is visible
          requestAnimationFrame(() => {
            const totalHeight =
              activeReasoningGroup.getBoundingClientRect().height
            reasoningContent.style.height = `${totalHeight}px`
          })
        } else {
          // Collapsing
          reasoningSection.classList.add('collapsed')
          reasoningContent.style.overflow = 'hidden'
          reasoningContent.style.height = '0'
        }
      }
    })

    // Update the streaming function to use activeReasoningGroup
    async function callPerplexityAPI(messages) {
      try {
        console.log('Model: ', trace.payload?.model)
        const response = await fetch(
          'https://api.perplexity.ai/chat/completions',
          {
            method: 'POST',
            headers: {
              Authorization: `Bearer ${trace.payload?.apiKey}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              model: trace.payload?.model || 'sonar-reasoning',
              messages: messages,
              return_images: false,
              return_related_questions: false,
              stream: true,
            }),
          }
        )

        const reader = response.body.getReader()
        const decoder = new TextDecoder()
        let buffer = ''
        let answer = ''
        let hasStartedAnswer = false
        let isFirstChunk = true
        let isInThinkBlock = false
        let thinkBuffer = ''
        let citations = []

        // Create and setup initial group
        const reasoningGroup = createReasoningGroup()
        reasoningContent.appendChild(reasoningGroup)
        activeReasoningGroup = reasoningGroup
        isStreaming = true

        // Add overflow handling styles
        reasoningContent.style.overflow = 'hidden'

        // Keep track of the current streaming step
        let currentStep = null

        // Update the createNewStep function
        function createNewStep(content, isComplete = false) {
          const step = createReasoningStep(
            content,
            activeReasoningGroup,
            activeReasoningGroup.children.length,
            citations
          )

          if (step) {
            // Add to DOM
            activeReasoningGroup.appendChild(step)

            // Show checkbox based on completion
            const checkbox = step.querySelector('.step-checkbox')
            if (checkbox) {
              checkbox.style.display = 'flex'
              if (isComplete) {
                checkbox.classList.add('is-checked')
                completedSteps.push(step)
              }
            }

            // Update height and ensure content is visible
            const totalHeight =
              activeReasoningGroup.getBoundingClientRect().height
            reasoningContent.style.height = `${totalHeight}px`

            // Force a reflow before scrolling
            void reasoningContent.offsetHeight

            // Attempt to scroll multiple times
            setTimeout(scrollToBottom, 10)
            setTimeout(scrollToBottom, 50)
            setTimeout(scrollToBottom, 100)

            return step
          }
          return null
        }

        // Update the updateStep function
        function updateStep(step, content, complete = false) {
          if (!step) return

          const contentDiv = step.querySelector('.step-content')
          if (contentDiv) {
            contentDiv.innerHTML = processCitations(content, citations, true)
          }

          if (complete && !completedSteps.includes(step)) {
            const checkbox = step.querySelector('.step-checkbox')
            if (checkbox) {
              checkbox.classList.add('is-checked')
            }
            completedSteps.push(step)
          }

          // Update height and ensure content is visible
          const totalHeight =
            activeReasoningGroup.getBoundingClientRect().height
          reasoningContent.style.height = `${totalHeight}px`

          // Force a reflow before scrolling
          void reasoningContent.offsetHeight

          // Attempt to scroll multiple times
          setTimeout(scrollToBottom, 10)
          setTimeout(scrollToBottom, 50)
          setTimeout(scrollToBottom, 100)
        }

        // Process thinking content
        function processThinkingContent(content) {
          thinkBuffer += content

          // Split buffer into complete and incomplete parts
          const parts = thinkBuffer.split('\n')
          const incompletePart = parts.pop() || ''
          const completeParts = parts.filter((part) => part.trim())

          // Process any complete lines
          for (const line of completeParts) {
            const trimmedLine = line.trim()
            if (trimmedLine) {
              // Complete the current step if it exists
              if (currentStep) {
                updateStep(currentStep, trimmedLine, true)
                currentStep = null
              } else {
                // Create a new completed step only if we don't have one for this content
                const existingSteps = Array.from(
                  activeReasoningGroup.querySelectorAll('.step-content')
                )
                const hasStep = existingSteps.some(
                  (step) =>
                    step.textContent.replace(/\s+/g, ' ').trim() === trimmedLine
                )
                if (!hasStep) {
                  createNewStep(trimmedLine, true)
                }
              }
            }
          }

          // Handle the remaining incomplete content
          if (incompletePart.trim()) {
            if (!currentStep) {
              // Create new streaming step
              currentStep = createNewStep(incompletePart.trim(), false)
            } else {
              // Update current streaming step
              updateStep(currentStep, incompletePart.trim(), false)
            }
          } else if (currentStep) {
            // If no incomplete part but we have a current step, complete it
            const contentDiv = currentStep.querySelector('.step-content')
            if (contentDiv) {
              const currentText = contentDiv.textContent
                .replace(/\s+/g, ' ')
                .trim()
              updateStep(currentStep, currentText, true)
            }
            currentStep = null
          }

          // Update buffer to only contain incomplete content
          thinkBuffer = incompletePart
        }

        // Process stream chunks
        function processStreamChunk(data) {
          if (data.choices[0].delta) {
            const { content } = data.choices[0].delta

            // Show container and answer section for non-reasoning models when streaming starts
            if (!isReasoningModel && container.style.display === 'none') {
              container.style.display = 'block'
              answerSection.classList.add('visible')
              answerSection.style.display = 'block'
            }

            // Update citations if present in the chunk
            if (data.citations) {
              citations = data.citations
              // Re-process all existing steps with new citations
              const steps =
                activeReasoningGroup.querySelectorAll('.reasoning-step')
              steps.forEach((step) => {
                const contentDiv = step.querySelector('.step-content')
                if (contentDiv) {
                  // Get original content without any HTML
                  const currentText = contentDiv.textContent
                    .replace(/\s+/g, ' ')
                    .trim()
                  // Process citations directly with isReasoning flag
                  contentDiv.innerHTML = processCitations(
                    currentText,
                    citations,
                    true
                  )
                }
              })
            }

            if (content !== null && content !== undefined) {
              // Handle think block content
              if (content.includes('<think>')) {
                isInThinkBlock = true
                const afterThink = content.split('<think>')[1] || ''
                if (afterThink) {
                  processThinkingContent(afterThink)
                }
              } else if (content.includes('</think>')) {
                isInThinkBlock = false
                const beforeThinkEnd = content.split('</think>')[0]
                if (beforeThinkEnd) {
                  processThinkingContent(beforeThinkEnd)
                }
                // Complete any remaining step
                if (currentStep) {
                  const contentDiv = currentStep.querySelector('.step-content')
                  if (contentDiv) {
                    const currentText = contentDiv.textContent
                    updateStep(currentStep, currentText, true)
                  }
                  currentStep = null
                }
                // Get content after </think>
                const afterThink = content.split('</think>')[1] || ''
                if (afterThink) {
                  answer += afterThink
                  updateAnswerContent(processCitations(answer, citations))
                }
              } else if (isInThinkBlock) {
                processThinkingContent(content)
              } else {
                // For non-think block content, just append directly to answer
                answer += content
                updateAnswerContent(processCitations(answer, citations))

                // Show answer section if this is the first content
                if (!hasStartedAnswer) {
                  hasStartedAnswer = true
                  setTimeout(() => {
                    reasoningSection.classList.add('collapsed')
                    reasoningSection.classList.add('has-answer')
                    const answerSection =
                      container.querySelector('.answer-section')
                    if (answerSection) {
                      answerSection.classList.add('visible')
                      answerSection.style.display = 'block'
                    }
                    scrollIntoViewSmooth(answerContent)
                  }, 1000)
                }
              }
            }
          }
        }

        while (true) {
          const { done, value } = await reader.read()
          if (done) {
            isStreaming = false
            break
          }

          if (isFirstChunk) {
            const reasoningIntro = container.querySelector('.reasoning-intro')
            if (reasoningIntro) {
              reasoningIntro.style.display = 'none'
            }
            isFirstChunk = false
          }

          const chunk = decoder.decode(value)
          let jsonBuffer = buffer + chunk
          buffer = ''

          // Split into lines and process each complete line
          const lines = jsonBuffer.split('\n')
          // Keep the last (potentially incomplete) line in buffer
          buffer = lines.pop() || ''

          for (const line of lines) {
            if (!line.trim() || !line.startsWith('data: ')) continue
            if (line === 'data: [DONE]') continue

            try {
              // Remove 'data: ' prefix and parse
              const jsonStr = line.slice(5)
              const data = JSON.parse(jsonStr)
              processStreamChunk(data)
            } catch (e) {
              // Skip incomplete chunks silently
            }
          }
        }

        // Process any remaining complete data in buffer
        if (buffer.trim() && buffer.startsWith('data: ')) {
          try {
            const data = JSON.parse(buffer.slice(5))
            processStreamChunk(data)
          } catch (e) {
            // Ignore parsing errors for final incomplete chunk
          }
        }
      } catch (error) {
        reasoningContent.textContent =
          'Error: Failed to get response from Perplexity API'
        const reasoningIntro = container.querySelector('.reasoning-intro')
        if (reasoningIntro) {
          reasoningIntro.style.display = 'none'
        }
      } finally {
        isStreaming = false
      }
    }

    // Function to find the closest scrollable parent
    function findScrollableParent(element) {
      let currentParent = element.parentElement
      while (currentParent) {
        const hasScrollableContent =
          currentParent.scrollHeight > currentParent.clientHeight
        const isScrollable =
          getComputedStyle(currentParent).overflow !== 'visible'

        if (hasScrollableContent && isScrollable) {
          return currentParent
        }
        currentParent = currentParent.parentElement
      }
      return null
    }

    // Function to smoothly scroll to an element
    function scrollIntoViewSmooth(element) {
      // Find the scrollable container by traversing up the DOM
      const scrollContainer = findScrollableParent(element)

      if (scrollContainer) {
        const elementRect = element.getBoundingClientRect()
        const containerRect = scrollContainer.getBoundingClientRect()
        const isElementInView =
          elementRect.top >= containerRect.top &&
          elementRect.bottom <= containerRect.bottom

        if (!isElementInView) {
          const scrollOffset = elementRect.bottom - containerRect.bottom + 50

          // Use scrollIntoView as a fallback if scrollTo is not working
          if (Math.abs(scrollContainer.scrollTop) < 1) {
            element.scrollIntoView({ behavior: 'smooth', block: 'end' })
          } else {
            scrollContainer.scrollTo({
              top: scrollContainer.scrollTop + scrollOffset,
              behavior: 'smooth',
            })
          }
        }
      } else {
        // If no scrollable container is found, try scrollIntoView as fallback
        element.scrollIntoView({ behavior: 'smooth', block: 'end' })
      }
    }

    // Function to convert markdown to HTML
    function markdownToHtml(markdown) {
      // First trim and clean the input
      let cleanedMarkdown = markdown
        .trim()
        .replace(/^\s+/gm, '') // Remove leading spaces from each line
        .replace(/\n{3,}/g, '\n\n') // Clean up excessive newlines

      // Process common introductory phrases
      cleanedMarkdown = cleanedMarkdown
        .replace(
          /^(here'?s?( is)?|let me|i will|i'll|allow me to|let's|okay,|so,|well,|now,|alright,).*?(:|->|\n)/i,
          ''
        )
        .replace(/^(here'?s?( is)?|let me|i will|i'll|allow me to|let's) /i, '')
        .replace(
          /^(a |the )?(breakdown|list|summary|overview) (of|for|about) /i,
          ''
        )
        .replace(/^(talking|speaking) about /i, '')
        .trim()

      // Split content into lines for better processing
      const lines = cleanedMarkdown.split('\n')
      const processedLines = []
      let currentList = null
      let isInParagraph = false
      let lastLineWasHeader = false

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim()
        const nextLine = i < lines.length - 1 ? lines[i + 1].trim() : ''
        const isCurrentLineHeader = line.startsWith('#')
        const isNextLineHeader = nextLine.startsWith('#')

        // Handle headers
        if (line.startsWith('###')) {
          if (currentList) {
            processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
            currentList = null
          }
          if (isInParagraph) {
            processedLines.push('</p>')
            isInParagraph = false
          }
          processedLines.push(
            `<h3 class="answer-h3">${line.slice(3).trim()}</h3>`
          )
          lastLineWasHeader = true
          continue
        }
        if (line.startsWith('##')) {
          if (currentList) {
            processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
            currentList = null
          }
          if (isInParagraph) {
            processedLines.push('</p>')
            isInParagraph = false
          }
          processedLines.push(
            `<h2 class="answer-h2">${line.slice(2).trim()}</h2>`
          )
          lastLineWasHeader = true
          continue
        }
        if (line.startsWith('#')) {
          if (currentList) {
            processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
            currentList = null
          }
          if (isInParagraph) {
            processedLines.push('</p>')
            isInParagraph = false
          }
          processedLines.push(
            `<h1 class="answer-h1">${line.slice(1).trim()}</h1>`
          )
          lastLineWasHeader = true
          continue
        }

        // Handle lists
        const orderedListMatch = line.match(/^\d+\.\s+(.+)/)
        const unorderedListMatch = line.match(/^-\s+(.+)/)

        if (orderedListMatch || unorderedListMatch) {
          if (isInParagraph) {
            processedLines.push('</p>')
            isInParagraph = false
          }

          const listContent = (orderedListMatch || unorderedListMatch)[1]
          const listType = orderedListMatch ? 'ol' : 'ul'

          if (!currentList) {
            currentList = { type: listType }
            // Add a newline before list only if previous line wasn't a header
            if (!lastLineWasHeader && processedLines.length > 0) {
              processedLines.push('')
            }
            processedLines.push(`<${listType} class="answer-list">`)
          } else if (currentList.type !== listType) {
            processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
            currentList = { type: listType }
            processedLines.push(`<${listType} class="answer-list">`)
          }

          processedLines.push(
            `<li class="answer-list-item">${listContent}</li>`
          )
          lastLineWasHeader = false
          continue
        }

        // Close list if we're not in a list item anymore
        if (currentList && line && !line.match(/^(-|\d+\.)\s/)) {
          processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
          currentList = null
        }

        // Handle paragraphs
        if (line) {
          if (!isInParagraph) {
            // Add a newline before paragraph only if previous line wasn't a header
            if (
              !lastLineWasHeader &&
              processedLines.length > 0 &&
              !isCurrentLineHeader
            ) {
              processedLines.push('')
            }
            processedLines.push('<p class="answer-paragraph">')
            isInParagraph = true
          }
          processedLines.push(line)
          lastLineWasHeader = false
        } else if (isInParagraph) {
          processedLines.push('</p>')
          isInParagraph = false
        }
      }

      // Close any open tags
      if (currentList) {
        processedLines.push(currentList.type === 'ol' ? '</ol>' : '</ul>')
      }
      if (isInParagraph) {
        processedLines.push('</p>')
      }

      // Join lines and apply remaining markdown formatting
      return (
        processedLines
          .join('\n')
          // Bold
          .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
          // Italic
          .replace(/\*(.*?)\*/g, '<em>$1</em>')
          // Code
          .replace(/`(.*?)`/g, '<code>$1</code>')
          // Links
          .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
          // Emojis (preserve as is)
          .replace(/[\u{1F300}-\u{1F9FF}]/gu, (match) => match)
          // Clean up any extra newlines between elements
          .replace(/>\n</g, '>\n<')
          .trim()
      )
    }

    // Update the answer content with markdown support
    function updateAnswerContent(text) {
      const trimmedText = text.trim()
      const formattedHtml = markdownToHtml(trimmedText)

      // Add styles for answer content
      const styles = `
        <style>
          .answer-content {
            font-size: 14px;
            line-height: 1.4;
            color: #374151;
            margin: 0;
            padding: 0;
          }
          .answer-paragraph {
            margin: 4px 0;
            padding: 0;
          }
          .answer-paragraph:first-child {
            margin-top: 0;
          }
          .answer-list {
            margin: 2px 0;
            padding-left: 20px;
          }
          .answer-list-item {
            margin: 2px 0;
            padding: 0;
          }
          .answer-h1 {
            font-size: 16px;
            margin: 12px 0 8px;
            font-weight: 600;
          }
          .answer-h1:first-child {
            margin-top: 0;
          }
          .answer-h2 {
            font-size: 14px;
            margin: 10px 0 6px;
            font-weight: 600;
          }
          .answer-h2:first-child {
            margin-top: 0;
          }
          .answer-h3 {
            font-size: 13px;
            margin: 8px 0 4px;
            font-weight: 600;
          }
          .answer-h3:first-child {
            margin-top: 0;
          }
          br {
            display: block;
            margin: 2px 0;
          }
          strong {
            font-weight: 600;
          }
        </style>
      `

      answerContent.innerHTML = styles + formattedHtml
    }

    // Function to create checkbox SVGs
    function createCheckbox() {
      const checkbox = document.createElement('div')
      checkbox.className = 'step-checkbox'
      checkbox.innerHTML = `
        <svg class="unchecked" viewBox="0 0 20 20" fill="none">
          <circle cx="10" cy="10" r="9" stroke="#9CA3AF" stroke-width="2"/>
        </svg>
        <svg class="checked" viewBox="0 0 20 20" fill="none">
          <circle cx="10" cy="10" r="10" fill="#111827"/>
          <path d="M14 7L8.5 12.5L6 10" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      `
      return checkbox
    }

    // Function to create a reasoning group
    function createReasoningGroup() {
      const group = document.createElement('div')
      group.className = 'reasoning-group'
      return group
    }

    // Function to process citations and convert [X] to links
    function processCitations(text, citations, isReasoning = false) {
      if (!citations || !citations.length) return text

      // First normalize spaces in the text
      let processedText = text.replace(/\s+/g, ' ').trim()

      // Add space before citations and convert to links in a single pass
      return processedText.replace(/\s*\[(\d+)\]/g, (match, num) => {
        const index = parseInt(num) - 1
        if (index >= 0 && index < citations.length) {
          // Add space only for answer section
          const prefix = !isReasoning ? ' ' : ''
          return (
            prefix +
            `<a href="${citations[index]}" target="_blank" class="citation-link">[${num}]</a>`
          )
        }
        return match
      })
    }

    // Function to create and append a new reasoning step
    function createReasoningStep(content, group, stepIndex, citations) {
      if (!content || content.length < 5 || content.trim().endsWith('?'))
        return null

      const step = document.createElement('div')
      step.className = 'reasoning-step'

      const checkbox = createCheckbox()
      checkbox.style.display = 'none' // Hide checkbox by default
      const contentDiv = document.createElement('div')
      contentDiv.className = 'step-content'

      // Process citations directly with isReasoning flag set to true
      contentDiv.innerHTML = processCitations(content, citations, true)

      step.appendChild(checkbox)
      step.appendChild(contentDiv)

      return step
    }

    // Add citation link styles
    const citationStyles = document.createElement('style')
    citationStyles.textContent = `
      .citation-link {
        color: #2563EB;
        text-decoration: none;
        font-weight: normal;
        cursor: pointer;
        margin: 0;
        padding: 0;
        display: inline;
      }
      .citation-link:hover {
        text-decoration: underline;
      }
      .reasoning-content .citation-link {
        color: #4B5563;
        font-weight: 500;
        display: inline;
        margin: 0;
        padding: 0;
      }
      .reasoning-content .citation-link:hover {
        color: #2563EB;
        text-decoration: underline;
      }
      .step-content {
        font-size: 12px;
        line-height: 1.4;
        padding-top: 1px;
        color: #4B5563;
      }
      .step-content a {
        color: inherit;
        margin: 0;
        padding: 0;
        display: inline;
      }
    `
    document.head.appendChild(citationStyles)

    // Create URL preview container
    const previewContainer = document.createElement('div')
    previewContainer.className = 'url-preview-container'
    document.body.appendChild(previewContainer)

    // Function to handle URL preview
    async function handleUrlPreview(event) {
      const link = event.target.closest('a')
      if (!link) return

      const url = link.href
      const rect = link.getBoundingClientRect()

      // Position the preview container
      const previewX = rect.left
      const previewY = rect.bottom + window.scrollY + 10 // 10px below the link

      previewContainer.style.left = `${previewX}px`
      previewContainer.style.top = `${previewY}px`

      try {
        // Show loading state with logo animation
        previewContainer.innerHTML = `
          <div class="url-preview-loading">
            <img src="/images/logo-anim.gif" alt="Loading preview..." />
          </div>
        `
        previewContainer.classList.add('visible')

        // Use mql to fetch URL preview
        const { data } = await window.mql(url, {
          data: {
            title: true,
            description: true,
            image: true,
          },
        })

        const { title, description, image } = data

        let previewHtml = '<div class="url-preview-content">'

        if (image?.url) {
          previewHtml += `<img class="url-preview-image" src="${
            image.url
          }" alt="${title || 'Preview'}" />`
        }

        if (title) {
          previewHtml += `<div class="url-preview-title">${title}</div>`
        }

        if (description) {
          previewHtml += `<div class="url-preview-description">${description}</div>`
        }

        previewHtml += '</div>'

        previewContainer.innerHTML = previewHtml
      } catch (error) {
        // If preview fails, show a simple preview with improved styling
        previewContainer.innerHTML = `
          <div class="url-preview-content">
            <div class="url-preview-title">${new URL(url).hostname}</div>
            <div class="url-preview-description">${url}</div>
          </div>
        `
      }
    }

    // Function to hide URL preview
    function hideUrlPreview() {
      previewContainer.classList.remove('visible')
    }

    // Add event listeners for URL preview
    element.addEventListener('mouseover', (e) => {
      const link = e.target.closest('a')
      if (link) {
        handleUrlPreview(e)
      }
    })

    element.addEventListener('mouseout', (e) => {
      const link = e.target.closest('a')
      if (link) {
        hideUrlPreview()
      }
    })

    if (trace.payload?.messages) {
      await callPerplexityAPI(trace.payload?.messages)
    }
    window.voiceflow.chat.interact({
      type: 'continue',
    })
  },
}
