/**
 * MailPilot – All-in-One Outlook Add-in (No Backend Required)
 *
 * Reads email via Office.js, calls Gemini API directly from the browser,
 * and renders AI analysis results. Zero server dependencies.
 */

// ──────────────────────────────────────
// Configuration
// ──────────────────────────────────────

const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta';
const GEMINI_MODEL = 'gemini-3.5-flash-lite';

const PRIORITY_LABELS = {
  1: 'Niski',
  2: 'Normalny',
  3: 'Ważny',
  4: 'Wysoki',
  5: 'Krytyczny',
};

const SYSTEM_PROMPT = `You are MailPilot, an expert AI email assistant. Your job is to analyze incoming emails and provide actionable intelligence.

RULES:
- Respond in the SAME LANGUAGE as the email (if email is in Polish, all your outputs should be in Polish, etc.)
- For draft replies: match the formality level of the original email. If it's casual, be casual. If it's formal, be formal.
- For draft replies: keep them concise and professional. Do NOT make up information you don't have.
- For priority scoring:
  1 = FYI, no action needed (notifications, confirmations)
  2 = Normal correspondence, can wait
  3 = Important, should handle today
  4 = High priority, needs quick response (deadline approaching, client request)
  5 = Critical/urgent (overdue payment, emergency, escalation)
- Be accurate with date extraction. Convert relative dates to absolute when possible.
- If the email doesn't require a reply (newsletter, notification, spam), leave draft_reply empty.
- For spam/marketing detection: be smart - a promotional email from a vendor you work with is "marketing", random unsolicited stuff is "spam".`;

// Gemini REST API response schema
const ANALYSIS_SCHEMA = {
  type: 'OBJECT',
  properties: {
    category: {
      type: 'STRING',
      description: 'Email category: Work, Finance, Personal, Support, Newsletter, Social, Shopping, Travel, or custom',
    },
    category_icon: {
      type: 'STRING',
      description: 'Single emoji icon for the category',
    },
    priority: {
      type: 'INTEGER',
      description: 'Priority 1-5',
    },
    urgency_reason: {
      type: 'STRING',
      description: 'Why urgent (if priority >= 4), empty otherwise',
    },
    summary: {
      type: 'STRING',
      description: '1-2 sentence summary',
    },
    draft_reply: {
      type: 'STRING',
      description: 'Reply draft in email language, empty if no reply needed',
    },
    is_spam_or_marketing: {
      type: 'BOOLEAN',
      description: 'True if spam/marketing',
    },
    spam_category: {
      type: 'STRING',
      description: 'spam/marketing/newsletter/promotional or empty',
    },
    extracted_dates: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          date: { type: 'STRING', description: 'Date in YYYY-MM-DD or as mentioned' },
          context: { type: 'STRING', description: 'What this date refers to' },
        },
        required: ['date', 'context'],
      },
    },
    extracted_amounts: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          value: { type: 'STRING', description: 'Numeric value' },
          currency: { type: 'STRING', description: 'Currency code' },
          context: { type: 'STRING', description: 'What this amount refers to' },
        },
        required: ['value', 'currency', 'context'],
      },
    },
    extracted_tasks: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          task: { type: 'STRING', description: 'Task description' },
          deadline: { type: 'STRING', description: 'Deadline if mentioned, empty otherwise' },
        },
        required: ['task', 'deadline'],
      },
    },
  },
  required: [
    'category', 'category_icon', 'priority', 'urgency_reason',
    'summary', 'draft_reply', 'is_spam_or_marketing', 'spam_category',
    'extracted_dates', 'extracted_amounts', 'extracted_tasks',
  ],
};

// ──────────────────────────────────────
// API Key Management (localStorage)
// ──────────────────────────────────────

function getApiKey() {
  return localStorage.getItem('mailpilot_api_key') || '';
}

function setApiKey(key) {
  localStorage.setItem('mailpilot_api_key', key.trim());
}

function hasApiKey() {
  return getApiKey().length > 10;
}

// ──────────────────────────────────────
// State management
// ──────────────────────────────────────

function showState(stateId) {
  document.querySelectorAll('.state').forEach((el) => el.classList.add('hidden'));
  const el = document.getElementById(stateId);
  if (el) el.classList.remove('hidden');
}

// ──────────────────────────────────────
// Office.js initialization
// ──────────────────────────────────────

let isOutlookContext = false;

if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      isOutlookContext = true;
      initSettingsUI();
      if (hasApiKey()) {
        analyzeCurrentEmail();
      } else {
        showState('settings-state');
      }
    } else {
      initSettingsUI();
      if (hasApiKey()) {
        initDemoMode();
      } else {
        showState('settings-state');
      }
    }
  });
} else {
  document.addEventListener('DOMContentLoaded', () => {
    initSettingsUI();
    if (hasApiKey()) {
      initDemoMode();
    } else {
      showState('settings-state');
    }
  });
}

// ──────────────────────────────────────
// Settings UI
// ──────────────────────────────────────

function initSettingsUI() {
  // Save key button
  const btnSave = document.getElementById('btn-save-key');
  const inputKey = document.getElementById('input-api-key');

  // Pre-fill if key exists
  if (hasApiKey()) {
    inputKey.value = getApiKey();
  }

  btnSave.addEventListener('click', () => {
    const key = inputKey.value.trim();
    if (!key || key.length < 10) {
      showToast('Podaj poprawny klucz API', 'error');
      return;
    }
    setApiKey(key);
    showToast('Klucz API zapisany!', 'success');

    // Start analysis or demo
    if (isOutlookContext) {
      analyzeCurrentEmail();
    } else {
      initDemoMode();
    }
  });

  // Toggle key visibility
  const btnToggle = document.getElementById('btn-toggle-key');
  btnToggle.addEventListener('click', () => {
    inputKey.type = inputKey.type === 'password' ? 'text' : 'password';
    btnToggle.textContent = inputKey.type === 'password' ? '👁️' : '🙈';
  });

  // Settings gear toggle
  const btnSettings = document.getElementById('btn-settings-toggle');
  btnSettings.addEventListener('click', () => {
    const settingsState = document.getElementById('settings-state');
    if (settingsState.classList.contains('hidden')) {
      // Fill current key
      inputKey.value = getApiKey();
      showState('settings-state');
    } else {
      // Go back
      if (isOutlookContext) {
        analyzeCurrentEmail();
      } else {
        initDemoMode();
      }
    }
  });

  // Enter key to save
  inputKey.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') btnSave.click();
  });
}

// ──────────────────────────────────────
// Read email from Outlook (Office.js)
// ──────────────────────────────────────

async function readCurrentEmail() {
  const item = Office.context.mailbox.item;
  if (!item) throw new Error('Nie można odczytać emaila');

  const subject = item.subject || '';
  const sender = item.from
    ? `${item.from.displayName || ''} <${item.from.emailAddress || ''}>`
    : '';
  const recipients = (item.to || []).map(
    (r) => r.emailAddress || r.displayName || ''
  );
  const attachmentNames = (item.attachments || []).map((a) => a.name || '');

  const body = await new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || 'Nie udało się odczytać treści emaila'));
      }
    });
  });

  return { subject, body, sender, recipients, attachmentNames };
}

// ──────────────────────────────────────
// Gemini API – Direct REST call (NO BACKEND!)
// ──────────────────────────────────────

async function callGeminiAPI(emailData) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error('Brak klucza API. Przejdź do ustawień.');

  const emailText = formatEmailForPrompt(emailData);

  const response = await fetch(
    `${GEMINI_API_BASE}/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        system_instruction: {
          parts: [{ text: SYSTEM_PROMPT }],
        },
        contents: [
          {
            role: 'user',
            parts: [{ text: emailText }],
          },
        ],
        generationConfig: {
          temperature: 0.3,
          maxOutputTokens: 2048,
          responseMimeType: 'application/json',
          responseSchema: ANALYSIS_SCHEMA,
        },
      }),
    }
  );

  const data = await response.json();

  // Handle API errors
  if (data.error) {
    const msg = data.error.message || 'Nieznany błąd API';
    if (msg.includes('API key')) {
      throw new Error('Nieprawidłowy klucz API. Sprawdź ustawienia.');
    }
    throw new Error(msg);
  }

  // Extract JSON from response
  if (!data.candidates || !data.candidates[0]?.content?.parts?.[0]?.text) {
    throw new Error('Gemini nie zwrócił odpowiedzi. Spróbuj ponownie.');
  }

  const resultText = data.candidates[0].content.parts[0].text;
  const result = JSON.parse(resultText);

  // Clamp priority
  result.priority = Math.max(1, Math.min(5, result.priority || 2));

  return result;
}

function formatEmailForPrompt(email) {
  const parts = [
    `Subject: ${email.subject || '(no subject)'}`,
    `From: ${email.sender || 'unknown'}`,
    `To: ${(email.recipients || []).join(', ') || 'unknown'}`,
  ];

  if (email.attachmentNames && email.attachmentNames.length > 0) {
    parts.push(`Attachments: ${email.attachmentNames.join(', ')}`);
  }

  parts.push(`\n--- Email Body ---\n${email.body || '(empty body)'}\n--- End of Email ---`);

  return parts.join('\n');
}

// ──────────────────────────────────────
// Main analysis flow
// ──────────────────────────────────────

let currentAnalysis = null;
let currentEmailData = null;

async function analyzeCurrentEmail() {
  showState('loading-state');
  try {
    currentEmailData = await readCurrentEmail();
    currentAnalysis = await callGeminiAPI(currentEmailData);
    renderResults(currentAnalysis);
    showState('results-state');
  } catch (err) {
    showError(err.message);
  }
}

async function analyzeEmailData(emailData) {
  showState('loading-state');
  try {
    currentEmailData = emailData;
    currentAnalysis = await callGeminiAPI(emailData);
    renderResults(currentAnalysis);
    showState('results-state');
  } catch (err) {
    showError(err.message);
  }
}

// ──────────────────────────────────────
// Render results
// ──────────────────────────────────────

function renderResults(data) {
  // Category
  document.getElementById('category-icon').textContent = data.category_icon || '📧';
  document.getElementById('category-name').textContent = data.category || 'Inne';

  // Priority
  const priority = data.priority || 2;
  const priorityBadge = document.getElementById('priority-badge');
  priorityBadge.className = `priority-badge priority-${priority}`;
  document.getElementById('priority-label').textContent =
    PRIORITY_LABELS[priority] || 'Normalny';

  // Summary
  document.getElementById('summary-text').textContent = data.summary || '';

  // Urgency
  const urgencyBanner = document.getElementById('urgency-banner');
  if (priority >= 4 && data.urgency_reason) {
    document.getElementById('urgency-text').textContent = data.urgency_reason;
    urgencyBanner.classList.remove('hidden');
  } else {
    urgencyBanner.classList.add('hidden');
  }

  // Spam banner
  const spamBanner = document.getElementById('spam-banner');
  if (data.is_spam_or_marketing) {
    document.getElementById('spam-type').textContent = data.spam_category || 'Marketing';
    spamBanner.classList.remove('hidden');
  } else {
    spamBanner.classList.add('hidden');
  }

  // Draft reply
  const draftCard = document.getElementById('card-draft');
  if (data.draft_reply) {
    document.getElementById('draft-text').value = data.draft_reply;
    draftCard.classList.remove('hidden');
  } else {
    draftCard.classList.add('hidden');
  }

  // Extracted data
  const hasExtracted = renderExtractedData(data);
  const extractedCard = document.getElementById('card-extracted');
  if (hasExtracted) {
    extractedCard.classList.remove('hidden');
  } else {
    extractedCard.classList.add('hidden');
  }

  bindActions();
}

function renderExtractedData(data) {
  let hasAny = false;

  // Dates
  const datesSection = document.getElementById('extracted-dates');
  const datesList = document.getElementById('dates-list');
  datesList.innerHTML = '';
  if (data.extracted_dates && data.extracted_dates.length > 0) {
    data.extracted_dates.forEach((d) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="data-value">${esc(d.date)}</span>
                       <span class="data-context">— ${esc(d.context)}</span>`;
      datesList.appendChild(li);
    });
    datesSection.classList.remove('hidden');
    hasAny = true;
  } else {
    datesSection.classList.add('hidden');
  }

  // Amounts
  const amountsSection = document.getElementById('extracted-amounts');
  const amountsList = document.getElementById('amounts-list');
  amountsList.innerHTML = '';
  if (data.extracted_amounts && data.extracted_amounts.length > 0) {
    data.extracted_amounts.forEach((a) => {
      const li = document.createElement('li');
      li.innerHTML = `<span class="data-value">${esc(a.value)} ${esc(a.currency)}</span>
                       <span class="data-context">— ${esc(a.context)}</span>`;
      amountsList.appendChild(li);
    });
    amountsSection.classList.remove('hidden');
    hasAny = true;
  } else {
    amountsSection.classList.add('hidden');
  }

  // Tasks
  const tasksSection = document.getElementById('extracted-tasks');
  const tasksList = document.getElementById('tasks-list');
  tasksList.innerHTML = '';
  if (data.extracted_tasks && data.extracted_tasks.length > 0) {
    data.extracted_tasks.forEach((t) => {
      const li = document.createElement('li');
      const dl = t.deadline ? ` (do: ${esc(t.deadline)})` : '';
      li.innerHTML = `<span class="data-value">•</span>
                       <span class="data-context">${esc(t.task)}${dl}</span>`;
      tasksList.appendChild(li);
    });
    tasksSection.classList.remove('hidden');
    hasAny = true;
  } else {
    tasksSection.classList.add('hidden');
  }

  return hasAny;
}

// ──────────────────────────────────────
// Actions
// ──────────────────────────────────────

function bindActions() {
  // Apply category
  document.getElementById('btn-apply-category').onclick = () => {
    if (isOutlookContext && currentAnalysis) {
      applyCategoryToEmail(currentAnalysis.category);
    } else {
      showToast('Dostępne tylko w Outlook', 'success');
    }
  };

  // Copy draft
  const btnCopy = document.getElementById('btn-copy-draft');
  if (btnCopy) {
    btnCopy.onclick = () => {
      const draft = document.getElementById('draft-text').value;
      if (draft) {
        navigator.clipboard
          .writeText(draft)
          .then(() => showToast('Skopiowano do schowka!', 'success'))
          .catch(() => showToast('Nie udało się skopiować', 'error'));
      }
    };
  }

  // Re-analyze
  document.getElementById('btn-reanalyze').onclick = () => {
    if (isOutlookContext) {
      analyzeCurrentEmail();
    } else if (currentEmailData) {
      analyzeEmailData(currentEmailData);
    }
  };
}

// ──────────────────────────────────────
// Outlook categories (Office.js)
// ──────────────────────────────────────

function applyCategoryToEmail(categoryName) {
  if (!categoryName) return;
  const name = `MailPilot: ${categoryName}`;

  Office.context.mailbox.masterCategories.getAsync((masterResult) => {
    if (masterResult.status !== Office.AsyncResultStatus.Succeeded) {
      showToast('Nie udało się pobrać kategorii', 'error');
      return;
    }

    const exists = masterResult.value.some((c) => c.displayName === name);

    const assign = () => {
      Office.context.mailbox.item.categories.addAsync([name], (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showToast(`Kategoria "${name}" zastosowana!`, 'success');
          const btn = document.getElementById('btn-apply-category');
          btn.innerHTML = '<span class="btn-icon">✅</span> Zastosowano!';
          btn.classList.replace('btn-primary', 'btn-success');
          setTimeout(() => {
            btn.innerHTML = '<span class="btn-icon">🏷️</span> Zastosuj kategorię';
            btn.classList.replace('btn-success', 'btn-primary');
          }, 2000);
        } else {
          showToast('Nie udało się przypisać kategorii', 'error');
        }
      });
    };

    if (!exists) {
      const color = Math.abs(hashCode(categoryName)) % 25;
      Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: name, color }],
        () => assign()
      );
    } else {
      assign();
    }
  });
}

// ──────────────────────────────────────
// Demo mode
// ──────────────────────────────────────

function initDemoMode() {
  showState('demo-state');

  document.getElementById('demo-subject').value =
    'Faktura za usługi IT – sierpień 2026';
  document.getElementById('demo-sender').value = 'jan.kowalski@techfirma.pl';
  document.getElementById('demo-body').value =
    'Dzień dobry,\n\nW załączeniu przesyłam fakturę nr FV/2026/08/142 na kwotę 12 500,00 PLN netto za usługi IT świadczone w sierpniu 2026.\n\nTermin płatności: 25 sierpnia 2026.\n\nProszę o potwierdzenie otrzymania.\n\nZ poważaniem,\nJan Kowalski\nTechFirma Sp. z o.o.';

  document.getElementById('btn-demo-analyze').onclick = () => {
    const emailData = {
      subject: document.getElementById('demo-subject').value,
      body: document.getElementById('demo-body').value,
      sender: document.getElementById('demo-sender').value,
      recipients: ['user@company.pl'],
      attachmentNames: [],
    };

    if (!emailData.subject && !emailData.body) {
      showToast('Wpisz temat lub treść emaila', 'error');
      return;
    }

    analyzeEmailData(emailData);
  };
}

// ──────────────────────────────────────
// Error handling
// ──────────────────────────────────────

function showError(message) {
  document.getElementById('error-message').textContent = message;
  showState('error-state');

  document.getElementById('btn-retry').onclick = () => {
    if (!hasApiKey()) {
      showState('settings-state');
    } else if (isOutlookContext) {
      analyzeCurrentEmail();
    } else {
      showState('demo-state');
    }
  };
}

// ──────────────────────────────────────
// Toast notifications
// ──────────────────────────────────────

function showToast(message, type = 'success') {
  const existing = document.querySelector('.toast');
  if (existing) existing.remove();

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  requestAnimationFrame(() => toast.classList.add('show'));

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 2500);
}

// ──────────────────────────────────────
// Utilities
// ──────────────────────────────────────

function esc(str) {
  const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  return String(str || '').replace(/[&<>"']/g, (m) => map[m]);
}

function hashCode(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = (hash << 5) - hash + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}
