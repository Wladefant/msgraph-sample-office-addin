// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* global location, Office, $ */

// <AuthUiSnippet>
// Handle to authentication pop dialog
/**
 * @type {Office.Dialog | undefined}
 */
let authDialog = undefined;

// @ts-ignore
var luxon = luxon || {
  DateTime: {
    local: () => {
      throw new Error('luxon not loaded');
    },
  },
};

// @ts-ignore
var OfficeRuntime = OfficeRuntime || {
  auth: {
    getAccessToken: () => {
      throw new Error('office.js not loaded');
    },
  },
};

// Build a base URL from the current location
function getBaseUrl() {
  return (
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '')
  );
}

// Process the response back from the auth dialog
/**
 * @param {{ message: string; origin: string | undefined; } | { error: number }} result
 */
function processConsent(result) {
  // @ts-ignore
  const message = JSON.parse(result.message);

  authDialog?.close();
  if (message.status === 'success') {
    showMainUi();
  } else {
    const error = JSON.stringify(
      message.result,
      Object.getOwnPropertyNames(message.result),
    );
    showStatus(`An error was returned from the consent dialog: ${error}`, true);
  }
}

// Use the Office Dialog API to show the interactive
// login UI
function showConsentPopup() {
  const authDialogUrl = `${getBaseUrl()}/consent.html`;

  Office.context.ui.displayDialogAsync(
    authDialogUrl,
    {
      height: 60,
      width: 30,
      promptBeforeOpen: false,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        authDialog = result.value;
        authDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processConsent,
        );
      } else {
        // Display error
        const error = JSON.stringify(
          result.error,
          Object.getOwnPropertyNames(result.error),
        );
        showStatus(`Could not open consent prompt dialog: ${error}`, true);
      }
    },
  );
}

// Inform the user we need to get their consent
function showConsentUi() {
  $('.container').empty();
  $('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'Consent for Microsoft Graph access needed',
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'In order to access your calendar, we need to get your permission to access the Microsoft Graph.',
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'We only need to do this once, unless you revoke your permission.',
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'Please click or tap the button below to give permission (opens a popup window).',
  }).appendTo('.container');
  $('<button/>', {
    class: 'primary-button',
    text: 'Give permission',
  })
    .on('click', showConsentPopup)
    .appendTo('.container');
}

// Display a status
/**
 * @param {unknown} message
 * @param {boolean} isError
 */
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`,
  })
    .append(
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: isError ? 'An error occurred' : 'Success',
      }),
    )
    .append(
      $('<p/>', {
        class: 'ms-fontSize-16 ms-fontWeight-regular',
        text: message,
      }),
    )
    .appendTo('.status');
}

/**
 * @param {boolean} show
 */
function toggleOverlay(show) {
  $('.overlay').css('display', show ? 'block' : 'none');
}
// </AuthUiSnippet>

// <MainUiSnippet>
function showMainUi() {
  $('.container').empty();

  // Use luxon to calculate the start
  // and end of the current week. Use
  // those dates to set the initial values
  // of the date pickers
  const now = luxon.DateTime.local();
  const startOfWeek = now.startOf('week');
  const endOfWeek = now.endOf('week');

  $('<h2/>', {
    class: 'ms-fontSize-24 ms-fontWeight-semibold',
    text: 'Select a date range to import',
  }).appendTo('.container');

  // Create the import form
  $('<form/>')
    .on('submit', getCalendar)
    .append(
      $('<label/>', {
        class: 'ms-fontSize-16 ms-fontWeight-semibold',
        text: 'Start',
      }),
    )
    .append(
      $('<input/>', {
        class: 'form-input',
        type: 'date',
        value: startOfWeek.toISODate(),
        id: 'viewStart',
      }),
    )
    .append(
      $('<label/>', {
        class: 'ms-fontSize-16 ms-fontWeight-semibold',
        text: 'End',
      }),
    )
    .append(
      $('<input/>', {
        class: 'form-input',
        type: 'date',
        value: endOfWeek.toISODate(),
        id: 'viewEnd',
      }),
    )
    .append(
      $('<input/>', {
        class: 'primary-button',
        type: 'submit',
        id: 'importButton',
        value: 'Import',
      }),
    )
    .appendTo('.container');

  $('<hr/>').appendTo('.container');

  $('<h2/>', {
    class: 'ms-fontSize-24 ms-fontWeight-semibold',
    text: 'Add event to calendar',
  }).appendTo('.container');

  // Create the new event form
  $('<form/>')
    .on('submit', createEvent)
    .append(
      $('<label/>', {
        class: 'ms-fontSize-16 ms-fontWeight-semibold',
        text: 'Subject',
      }),
    )
    .append(
      $('<input/>', {
        class: 'form-input',
        type: 'text',
        required: true,
        id: 'eventSubject',
      }),
    )
    .append(
      $('<label/>', {
        class: 'ms-fontSize-16 ms-fontWeight-semibold',
        text: 'Start',
      }),
    )
    .append(
      $('<input/>', {
        class: 'form-input',
        type: 'datetime-local',
        required: true,
        id: 'eventStart',
      }),
    )
    .append(
      $('<label/>', {
        class: 'ms-fontSize-16 ms-fontWeight-semibold',
        text: 'End',
      }),
    )
    .append(
      $('<input/>', {
        class: 'form-input',
        type: 'datetime-local',
        required: true,
        id: 'eventEnd',
      }),
    )
    .append(
      $('<input/>', {
        class: 'primary-button',
        type: 'submit',
        id: 'importButton',
        value: 'Create',
      }),
    )
    .appendTo('.container');
}
// </MainUiSnippet>

// <WriteToEmailSnippet>
/**
 * @param {any[]} events
 */
async function writeEventsToEmail(events) {
  let emailBody = 'Here are your calendar events:\n\n';

  events.forEach((event) => {
    emailBody += `Subject: ${event.subject}\n`;
    emailBody += `Organizer: ${event.organizer.emailAddress.name}\n`;
    emailBody += `Start: ${event.start.dateTime}\n`;
    emailBody += `End: ${event.end.dateTime}\n\n`;
  });

  Office.context.mailbox.item.body.setAsync(
    emailBody,
    { coercionType: Office.CoercionType.Text },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showStatus(`Error writing to email: ${asyncResult.error.message}`, true);
      } else {
        showStatus('Events written to email', false);
      }
    }
  );
}
// </WriteToEmailSnippet>

// <GetCalendarSnippet>
/**
 * @param {{ preventDefault: () => void; }} evt
 */
async function getCalendar(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  try {
    const apiToken = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
    });

    const viewStart = $('#viewStart').val();
    const viewEnd = $('#viewEnd').val();

    const requestUrl = `${getBaseUrl()}/graph/calendarview?viewStart=${viewStart}&viewEnd=${viewEnd}`;

    const response = await fetch(requestUrl, {
      headers: {
        authorization: `Bearer ${apiToken}`,
      },
    });

    if (response.ok) {
      const events = await response.json();
      if (events.length > 0) writeEventsToEmail(events);
      showStatus(`Imported ${events.length} events`, false);
    } else {
      const error = await response.json();
      showStatus(
        `Error getting events from calendar: ${JSON.stringify(error)}`,
        true,
      );
    }

    toggleOverlay(false);
  } catch (err) {
    console.log(`Error: ${JSON.stringify(err)}`);
    showStatus(
      `Exception getting events from calendar: ${JSON.stringify(err)}`,
      true,
    );
  }
}
// </GetCalendarSnippet>

// <CreateEventSnippet>
/**
 * @param {{ preventDefault: () => void; }} evt
 */
async function createEvent(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  const apiToken = await OfficeRuntime.auth.getAccessToken({
    allowSignInPrompt: true,
  });

  const payload = {
    eventSubject: $('#eventSubject').val(),
    eventStart: $('#eventStart').val(),
    eventEnd: $('#eventEnd').val(),
  };

  const requestUrl = `${getBaseUrl()}/graph/newevent`;

  const response = await fetch(requestUrl, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${apiToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (response.ok) {
    showStatus('Event created', false);
  } else {
    const error = await response.json();
    showStatus(`Error creating event: ${JSON.stringify(error)}`, true);
  }

  toggleOverlay(false);
}
// </CreateEventSnippet>

// <OfficeReadySnippet>
Office.onReady((info) => {
  // Only run if we're inside Outlook
  if (info.host === Office.HostType.Outlook) {
    $(async function () {
      let apiToken = '';
      try {
        apiToken = await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
        });
        console.log(`API Token: ${apiToken}`);
      } catch (error) {
        console.log(`getAccessToken error: ${JSON.stringify(error)}`);
        // Fall back to interactive login
        showConsentUi();
      }

      // Call auth status API to see if we need to get consent
      const authStatusResponse = await fetch(`${getBaseUrl()}/auth/status`, {
        headers: {
          authorization: `Bearer ${apiToken}`,
        },
      });

      const authStatus = await authStatusResponse.json();
      if (authStatus.status === 'consent_required') {
        showConsentUi();
      } else {
        // report error
        if (authStatus.status === 'error') {
          const error = JSON.stringify(
            authStatus.error,
            Object.getOwnPropertyNames(authStatus.error),
          );
          showStatus(`Error checking auth status: ${error}`, true);
        } else {
          showMainUi();
        }
      }
    });
  }
});
// </OfficeReadySnippet>
