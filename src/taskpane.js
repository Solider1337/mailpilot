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
// Settings, Theme & i18n
// ──────────────────────────────────────

const DICT = {
  pl: {
    settingsTitle: "Ustawienia", languageLabel: "Język aplikacji:", themeLabel: "Motyw:", themeLight: "Jasny", themeDark: "Ciemny", planLabel: "Subskrypcja:", managePlanBtn: "Zarządzaj",
    riskDetected: "Wykryto ryzyko:", summary: "Podsumowanie", keyTakeaways: "Kluczowe informacje", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Gotowa odpowiedź", tabCustom: "Własna odpowiedź", chooseStyle: "Wybierz styl odpowiedzi:", copy: "Kopiuj",
    instructionForAi: "Instrukcja dla AI:", generateReply: "Generuj odpowiedź", result: "Wynik:",
    actionItems: "Zadania do wykonania", extractedEntities: "Wykryte informacje", contacts: "Kontakty", datesEvents: "Daty / Wydarzenia", financials: "Finanse",
    applyCategory: "Zastosuj kategorię", forceAnalysis: "Wymuś analizę",
    orgPrefix: "Organizacja:", b2bPlanPrefix: "Subskrypcja firmowa:", planPrefix: "Plan:", active: "Aktywny", accountPrefix: "Konto:", costPrefix: "Koszt: 10$ / miesiąc"
  },
  en: {
    settingsTitle: "Settings", languageLabel: "App Language:", themeLabel: "Theme:", themeLight: "Light", themeDark: "Dark", planLabel: "Subscription:", managePlanBtn: "Manage",
    riskDetected: "Risk detected:", summary: "Summary", keyTakeaways: "Key Takeaways", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Preset Reply", tabCustom: "Custom Reply", chooseStyle: "Choose reply style:", copy: "Copy",
    instructionForAi: "Instruction for AI:", generateReply: "Generate Reply", result: "Result:",
    actionItems: "Action Items & Tasks", extractedEntities: "Extracted Entities", contacts: "Contacts", datesEvents: "Dates / Events", financials: "Financials",
    applyCategory: "Apply Category", forceAnalysis: "Force Analysis",
    orgPrefix: "Organization:", b2bPlanPrefix: "Corporate Subscription:", planPrefix: "Plan:", active: "Active", accountPrefix: "Account:", costPrefix: "Cost: $10 / month"
  },
  de: {
    settingsTitle: "Einstellungen", languageLabel: "Sprache:", themeLabel: "Thema:", themeLight: "Hell", themeDark: "Dunkel", planLabel: "Abonnement:", managePlanBtn: "Verwalten",
    riskDetected: "Risiko erkannt:", summary: "Zusammenfassung", keyTakeaways: "Wichtigste Punkte", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Vorgefertigte Antwort", tabCustom: "Eigene Antwort", chooseStyle: "Antwortstil wählen:", copy: "Kopieren",
    instructionForAi: "Anweisung für KI:", generateReply: "Antwort generieren", result: "Ergebnis:",
    actionItems: "Aufgaben", extractedEntities: "Extrahierte Entitäten", contacts: "Kontakte", datesEvents: "Daten / Ereignisse", financials: "Finanzen",
    applyCategory: "Kategorie anwenden", forceAnalysis: "Analyse erzwingen"
  },
  es: {
    settingsTitle: "Configuración", languageLabel: "Idioma:", themeLabel: "Tema:", themeLight: "Claro", themeDark: "Oscuro", planLabel: "Suscripción:", managePlanBtn: "Gestionar",
    riskDetected: "Riesgo detectado:", summary: "Resumen", keyTakeaways: "Puntos clave", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Respuesta predefinida", tabCustom: "Respuesta personalizada", chooseStyle: "Elige el estilo:", copy: "Copiar",
    instructionForAi: "Instrucción para IA:", generateReply: "Generar respuesta", result: "Resultado:",
    actionItems: "Tareas", extractedEntities: "Entidades extraídas", contacts: "Contactos", datesEvents: "Fechas / Eventos", financials: "Finanzas",
    applyCategory: "Aplicar categoría", forceAnalysis: "Forzar análisis"
  },
  fr: {
    settingsTitle: "Paramètres", languageLabel: "Langue:", themeLabel: "Thème:", themeLight: "Clair", themeDark: "Sombre", planLabel: "Abonnement:", managePlanBtn: "Gérer",
    riskDetected: "Risque détecté:", summary: "Résumé", keyTakeaways: "Points clés", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Réponse prédéfinie", tabCustom: "Réponse personnalisée", chooseStyle: "Choisir le style:", copy: "Copier",
    instructionForAi: "Instruction pour l'IA:", generateReply: "Générer la réponse", result: "Résultat:",
    actionItems: "Tâches", extractedEntities: "Entités extraites", contacts: "Contacts", datesEvents: "Dates / Événements", financials: "Finances",
    applyCategory: "Appliquer la catégorie", forceAnalysis: "Forcer l'analyse"
  },
  it: {
    settingsTitle: "Impostazioni", languageLabel: "Lingua:", themeLabel: "Tema:", themeLight: "Chiaro", themeDark: "Scuro", planLabel: "Abbonamento:", managePlanBtn: "Gestisci",
    riskDetected: "Rischio rilevato:", summary: "Riepilogo", keyTakeaways: "Punti chiave", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Risposta predefinita", tabCustom: "Risposta personalizzata", chooseStyle: "Scegli lo stile:", copy: "Copia",
    instructionForAi: "Istruzione per IA:", generateReply: "Genera risposta", result: "Risultato:",
    actionItems: "Attività", extractedEntities: "Entità estratte", contacts: "Contatti", datesEvents: "Date / Eventi", financials: "Finanze",
    applyCategory: "Applica categoria", forceAnalysis: "Forza analisi"
  },
  pt: {
    settingsTitle: "Configurações", languageLabel: "Idioma:", themeLabel: "Tema:", themeLight: "Claro", themeDark: "Escuro", planLabel: "Assinatura:", managePlanBtn: "Gerenciar",
    riskDetected: "Risco detectado:", summary: "Resumo", keyTakeaways: "Pontos chave", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Resposta predefinida", tabCustom: "Resposta personalizada", chooseStyle: "Escolher o estilo:", copy: "Copiar",
    instructionForAi: "Instrução para IA:", generateReply: "Gerar resposta", result: "Resultado:",
    actionItems: "Tarefas", extractedEntities: "Entidades extraídas", contacts: "Contatos", datesEvents: "Datas / Eventos", financials: "Finanças",
    applyCategory: "Aplicar categoria", forceAnalysis: "Forçar análise"
  },
  uk: {
    settingsTitle: "Налаштування", languageLabel: "Мова:", themeLabel: "Тема:", themeLight: "Світла", themeDark: "Темна", planLabel: "Підписка:", managePlanBtn: "Керувати",
    riskDetected: "Виявлено ризик:", summary: "Резюме", keyTakeaways: "Ключові моменти", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "Готова відповідь", tabCustom: "Власна відповідь", chooseStyle: "Виберіть стиль:", copy: "Копіювати",
    instructionForAi: "Інструкція для ШІ:", generateReply: "Згенерувати відповідь", result: "Результат:",
    actionItems: "Завдання", extractedEntities: "Витягнуті дані", contacts: "Контакти", datesEvents: "Дати / Події", financials: "Фінанси",
    applyCategory: "Застосувати категорію", forceAnalysis: "Примусовий аналіз"
  },
  zh: {
    settingsTitle: "设置", languageLabel: "语言:", themeLabel: "主题:", themeLight: "浅色", themeDark: "深色", planLabel: "订阅:", managePlanBtn: "管理",
    riskDetected: "检测到风险:", summary: "摘要", keyTakeaways: "关键信息", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "预设回复", tabCustom: "自定义回复", chooseStyle: "选择回复风格:", copy: "复制",
    instructionForAi: "AI指令:", generateReply: "生成回复", result: "结果:",
    actionItems: "任务", extractedEntities: "提取的实体", contacts: "联系人", datesEvents: "日期 / 事件", financials: "财务",
    applyCategory: "应用分类", forceAnalysis: "强制分析"
  },
  ja: {
    settingsTitle: "設定", languageLabel: "言語:", themeLabel: "テーマ:", themeLight: "ライト", themeDark: "ダーク", planLabel: "サブスクリプション:", managePlanBtn: "管理",
    riskDetected: "リスクを検出:", summary: "要約", keyTakeaways: "重要なポイント", mailMaestroMode: "✨ Mail Maestro Mode",
    tabPreset: "定型返信", tabCustom: "カスタム返信", chooseStyle: "返信スタイルを選択:", copy: "コピー",
    instructionForAi: "AIへの指示:", generateReply: "返信を生成", result: "結果:",
    actionItems: "タスク", extractedEntities: "抽出された情報", contacts: "連絡先", datesEvents: "日付 / イベント", financials: "財務",
    applyCategory: "カテゴリを適用", forceAnalysis: "分析を強制"
  }
};

let currentLang = localStorage.getItem('mailpilot-lang') || navigator.language.split('-')[0];
if (!DICT[currentLang]) currentLang = 'en';

function applyLanguage(lang) {
  currentLang = lang;
  localStorage.setItem('mailpilot-lang', lang);
  const dict = DICT[lang] || DICT.en;
  
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.getAttribute('data-i18n');
    if (dict[key]) {
      el.textContent = dict[key];
    }
  });
}

function applyTheme(theme) {
  localStorage.setItem('mailpilot-theme', theme);
  if (theme === 'dark') {
    document.body.classList.add('dark-theme');
  } else {
    document.body.classList.remove('dark-theme');
  }
}

function initSettings() {
  const btnSettings = document.getElementById('btn-settings');
  const btnCloseSettings = document.getElementById('btn-close-settings');
  const modal = document.getElementById('settings-modal');
  const selLang = document.getElementById('settings-language');
  const selTheme = document.getElementById('settings-theme');
  
  // Load saved prefs
  applyLanguage(currentLang);
  selLang.value = currentLang;
  
  const savedTheme = localStorage.getItem('mailpilot-theme') || 'light';
  applyTheme(savedTheme);
  selTheme.value = savedTheme;
  
  btnSettings.onclick = () => modal.classList.remove('hidden');
  btnCloseSettings.onclick = () => modal.classList.add('hidden');
  
  selLang.onchange = (e) => {
    applyLanguage(e.target.value);
    if (currentEmailData) {
      analyzeEmailData(currentEmailData);
    }
  };
  selTheme.onchange = (e) => applyTheme(e.target.value);
  
  document.getElementById('btn-manage-plan').onclick = () => {
    window.open('https://autome.github.io/mailpilot/checkout', '_blank');
  };
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

      initSettings();
      analyzeCurrentEmail();
    } else {
      initDemoMode();
    }
  });
} else {
  // Test fallback
  window.onload = () => {
    initSettings();
    showState('loading-state');
    setTimeout(() => {
      analyzeEmailData(MOCK_EMAIL);
    }, 1000);
  };
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
      userEmail: userEmail,
      language: currentLang
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

  // Ustaw odznakę organizacji i status planu w Ustawieniach
  if (data._auth_info) {
    const badge = document.getElementById('org-badge');
    const planInfo = document.getElementById('plan-info');
    const btnManagePlan = document.getElementById('btn-manage-plan');
    const dict = DICT[currentLang] || DICT.en;
    
    if (data._auth_info.reason === 'B2B User') {
      const orgName = data._auth_info.organization || 'B2B';
      badge.textContent = `${dict.orgPrefix || DICT.en.orgPrefix} ${orgName}`;
      badge.classList.remove('hidden');
      planInfo.textContent = `${dict.b2bPlanPrefix || DICT.en.b2bPlanPrefix} ${orgName}`;
      btnManagePlan.classList.add('hidden');
    } else {
      badge.classList.add('hidden');
      planInfo.innerHTML = `${dict.planPrefix || DICT.en.planPrefix} <strong style="color:var(--success)">${dict.active || DICT.en.active}</strong><br>${dict.accountPrefix || DICT.en.accountPrefix} ${userEmail}<br>${dict.costPrefix || DICT.en.costPrefix}`;
      btnManagePlan.classList.remove('hidden');
    }
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
            mode: "preset",
            language: currentLang
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
          mode: 'generate',
          language: currentLang
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
