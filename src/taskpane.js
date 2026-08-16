/**
 * MailPilot – All-in-One Outlook Add-in (No Backend Required)
 *
 * Reads email via Office.js, calls Gemini API directly from the browser,
 * and renders AI analysis results. Zero server dependencies.
 */

const BACKEND_URL = 'https://backend-beta-one-53.vercel.app/api/analyze';

const PRIORITY_LABELS = {
  1: 'Niski',
  2: 'Normalny',
  3: 'Ważny',
  4: 'Wysoki',
  5: 'Krytyczny',
};

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
let userEmail = 'demo@example.com';

if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      isOutlookContext = true;
      userEmail = Office.context.mailbox.userProfile.emailAddress;
      
      // Nasłuchuj zmiany zaznaczonego maila
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        function() {
          showState('loading-state');
          setTimeout(() => {
            analyzeCurrentEmail();
          }, 500);
        },
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showError("Błąd przypinania: " + asyncResult.error.message);
          }
        }
      );

      // Zabezpieczenie (Fallback): Wymuszone sprawdzanie zmiany maila co 1 sekundę
      let lastItemId = Office.context.mailbox.item ? Office.context.mailbox.item.itemId : null;
      setInterval(() => {
        const currentItem = Office.context.mailbox.item;
        if (currentItem && currentItem.itemId !== lastItemId) {
          lastItemId = currentItem.itemId;
          console.log("Wykryto zmianę maila przez polling!");
          showState('loading-state');
          setTimeout(() => {
            analyzeCurrentEmail();
          }, 500);
        }
      }, 1000);

      analyzeCurrentEmail();
    } else {
      initDemoMode();
    }
  });
} else {
  document.addEventListener('DOMContentLoaded', () => {
    initDemoMode();
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
// Backend API (Vercel/Local)
// ──────────────────────────────────────

async function callBackendAPI(emailData) {
  const emailText = formatEmailForPrompt(emailData);

  const response = await fetch(BACKEND_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      emailText: emailText,
      userEmail: userEmail
    })
  });

  const data = await response.json();

  if (!response.ok) {
    if (response.status === 403) {
      document.getElementById('current-user-email').textContent = userEmail;
      showState('license-state');
      throw new Error('LICENSE_ERROR');
    }
    throw new Error(data.detail || 'Wystąpił błąd połączenia z serwerem');
  }

  // Ustaw odznakę organizacji z backendu
  if (data._auth_info && data._auth_info.organization) {
    const badge = document.getElementById('org-badge');
    badge.textContent = `Organizacja: ${data._auth_info.organization}`;
    badge.classList.remove('hidden');
  }

  // Wymuszony zakres priority
  data.priority = Math.max(1, Math.min(5, data.priority || 2));

  return data;
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
    currentAnalysis = await callBackendAPI(currentEmailData);
    renderResults(currentAnalysis);
    showState('results-state');
  } catch (err) {
    if (err.message !== 'LICENSE_ERROR') showError(err.message);
  }
}

async function analyzeEmailData(emailData) {
  showState('loading-state');
  try {
    currentEmailData = emailData;
    currentAnalysis = await callBackendAPI(emailData);
    renderResults(currentAnalysis);
    showState('results-state');
  } catch (err) {
    if (err.message !== 'LICENSE_ERROR') showError(err.message);
  }
}

// ──────────────────────────────────────
// Render results
// ──────────────────────────────────────

function renderResults(data) {
  // Summary
  document.getElementById('summary-text').textContent = data.summary || '';

  // Spam banner
  const spamBanner = document.getElementById('spam-banner');
  if (data.is_spam_or_marketing) {
    document.getElementById('spam-type').textContent = data.spam_category || 'Marketing';
    spamBanner.classList.remove('hidden');
  } else {
    spamBanner.classList.add('hidden');
  }

  // Red Flags
  const redflagsBanner = document.getElementById('redflags-banner');
  const redflagsList = document.getElementById('redflags-list');
  redflagsList.innerHTML = '';
  if (data.red_flags && data.red_flags.length > 0) {
    data.red_flags.forEach((flag) => {
      const li = document.createElement('li');
      li.textContent = flag;
      redflagsList.appendChild(li);
    });
    redflagsBanner.classList.remove('hidden');
  } else {
    redflagsBanner.classList.add('hidden');
  }

  // Mail Maestro Draft & Improve (Hybrid Mode)
  const draftCard = document.getElementById('card-draft');
  const tabPreset = document.getElementById('tab-preset');
  const tabCustom = document.getElementById('tab-custom');
  
  const presetArea = document.getElementById('draft-preset-area');
  const customArea = document.getElementById('draft-custom-area');
  
  const draftSelector = document.getElementById('draft-selector');
  const draftPresetText = document.getElementById('draft-preset-text');
  
  const draftPrompt = document.getElementById('draft-prompt');
  const btnGenerateDraft = document.getElementById('btn-generate-draft');
  const draftCustomResultArea = document.getElementById('draft-custom-result-area');
  const draftCustomText = document.getElementById('draft-custom-text');

  // Setup tabs
  tabPreset.onclick = () => {
    tabPreset.className = 'btn btn-primary';
    tabCustom.className = 'btn btn-secondary';
    presetArea.classList.remove('hidden');
    customArea.classList.add('hidden');
  };

  tabCustom.onclick = () => {
    tabPreset.className = 'btn btn-secondary';
    tabCustom.className = 'btn btn-primary';
    presetArea.classList.add('hidden');
    customArea.classList.remove('hidden');
  };

  // Populate Preset options from initial API payload
  draftSelector.innerHTML = '';
  if (data.draft_types && data.draft_types.length > 0) {
    data.draft_types.forEach((type) => {
      const option = document.createElement('option');
      option.value = type;
      option.textContent = type;
      draftSelector.appendChild(option);
    });
    draftPresetText.value = data.draft_reply || '';
    
    // Changing the preset triggers a generate call with mode "preset"
    draftSelector.onchange = async (e) => {
      const selectedType = e.target.value;
      draftPresetText.value = "Generuję nową odpowiedź...";
      draftPresetText.disabled = true;
      try {
        const response = await fetch(BACKEND_URL.replace('/api/analyze', '/api/draft'), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            emailText: formatEmailForPrompt(currentEmailData),
            userEmail: userEmail,
            prompt: selectedType,
            mode: "preset"
          })
        });
        const resData = await response.json();
        if (resData.draft) {
          draftPresetText.value = resData.draft;
        }
      } catch (err) {
        draftPresetText.value = "Wystąpił błąd podczas generowania.";
      }
      draftPresetText.disabled = false;
    };
  }

  // Custom Generate action
  btnGenerateDraft.onclick = async () => {
    const promptValue = draftPrompt.value.trim();
    if (!promptValue) {
      showToast('Wpisz instrukcję dla AI!', 'error');
      return;
    }

    const btnIcon = btnGenerateDraft.querySelector('.btn-icon');
    const originalContent = btnGenerateDraft.innerHTML;
    btnGenerateDraft.innerHTML = 'Generowanie...';
    btnGenerateDraft.disabled = true;
    
    try {
      const response = await fetch(BACKEND_URL.replace('/api/analyze', '/api/draft'), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emailText: formatEmailForPrompt(currentEmailData),
          userEmail: userEmail,
          prompt: promptValue,
          mode: 'generate'
        })
      });
      const resData = await response.json();
      
      if (resData.draft) {
        draftCustomText.value = resData.draft;
        draftCustomResultArea.classList.remove('hidden');
      } else {
        showToast('Błąd generowania', 'error');
      }
    } catch (err) {
      showToast('Wystąpił błąd komunikacji', 'error');
    }
    
    btnGenerateDraft.innerHTML = originalContent;
    btnGenerateDraft.disabled = false;
  };

  draftCard.classList.remove('hidden');
  
  // Wyczyść wyniki z poprzedniego maila dla zakładki custom
  draftCustomResultArea.classList.add('hidden');
  draftCustomText.value = '';
  draftPrompt.value = '';

  // Extracted data
  renderExtractedData(data);

  bindActions();
}

function renderExtractedData(data) {
  let hasEntities = false;
  let hasTasks = false;

  // Dates
  const datesSection = document.getElementById('extracted-dates');
  const datesList = document.getElementById('dates-list');
  datesList.innerHTML = '';
  if (data.extracted_dates && data.extracted_dates.length > 0) {
    data.extracted_dates.forEach((d) => {
      const div = document.createElement('div');
      div.className = 'entity-card';
      div.innerHTML = `<strong>${esc(d.date)}</strong> - ${esc(d.context)}`;
      datesList.appendChild(div);
    });
    datesSection.classList.remove('hidden');
    hasEntities = true;
  } else {
    datesSection.classList.add('hidden');
  }

  // Amounts
  const amountsSection = document.getElementById('extracted-amounts');
  const amountsList = document.getElementById('amounts-list');
  amountsList.innerHTML = '';
  if (data.extracted_amounts && data.extracted_amounts.length > 0) {
    data.extracted_amounts.forEach((a) => {
      const div = document.createElement('div');
      div.className = 'entity-card';
      div.innerHTML = `<strong>${esc(a.value)} ${esc(a.currency)}</strong> - ${esc(a.context)}`;
      amountsList.appendChild(div);
    });
    amountsSection.classList.remove('hidden');
    hasEntities = true;
  } else {
    amountsSection.classList.add('hidden');
  }

  // Tasks
  const tasksSection = document.getElementById('section-tasks');
  const tasksGrid = document.getElementById('tasks-grid');
  tasksGrid.innerHTML = '';
  if (data.extracted_tasks && data.extracted_tasks.length > 0) {
    data.extracted_tasks.forEach((t, i) => {
      const div = document.createElement('div');
      div.className = 'task-card';
      div.innerHTML = `
        <div class="task-header">
          <span class="task-label">Task ${i + 1}</span>
        </div>
        <div class="task-title">${esc(t.task)}</div>
        ${t.deadline ? `<div class="task-time">${esc(t.deadline)}</div>` : ''}
      `;
      tasksGrid.appendChild(div);
    });
    tasksSection.classList.remove('hidden');
    hasTasks = true;
  } else {
    tasksSection.classList.add('hidden');
  }

  // Contacts
  const contactsSection = document.getElementById('extracted-contacts');
  const contactsGrid = document.getElementById('contacts-grid');
  if (contactsGrid) {
    contactsGrid.innerHTML = '';
    if (data.extracted_contacts && data.extracted_contacts.length > 0) {
      data.extracted_contacts.forEach((c) => {
        const div = document.createElement('div');
        div.className = 'contact-card';
        const initial = c.name ? c.name.charAt(0).toUpperCase() : '?';
        div.innerHTML = `
          <div class="contact-avatar">${initial}</div>
          <div class="contact-info">
            <div class="contact-name">${esc(c.name)}</div>
            <div class="contact-role">${esc(c.info)}</div>
          </div>
        `;
        contactsGrid.appendChild(div);
      });
      contactsSection.classList.remove('hidden');
      hasEntities = true;
    } else {
      contactsSection.classList.add('hidden');
    }
  }

  const entitiesSection = document.getElementById('section-entities');
  if (hasEntities) {
    entitiesSection.classList.remove('hidden');
  } else {
    entitiesSection.classList.add('hidden');
  }
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

  // Copy preset draft
  const btnCopyPreset = document.getElementById('btn-copy-preset');
  if (btnCopyPreset) {
    btnCopyPreset.onclick = () => {
      const draft = document.getElementById('draft-preset-text').value;
      if (draft) {
        navigator.clipboard
          .writeText(draft)
          .then(() => showToast('Skopiowano do schowka!', 'success'))
          .catch(() => showToast('Nie udało się skopiować', 'error'));
      }
    };
  }

  // Copy custom draft
  const btnCopyCustom = document.getElementById('btn-copy-custom');
  if (btnCopyCustom) {
    btnCopyCustom.onclick = () => {
      const draft = document.getElementById('draft-custom-text').value;
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
    if (isOutlookContext) {
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
