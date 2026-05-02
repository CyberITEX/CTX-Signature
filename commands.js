// Event-based handler — runs in JS-only runtime (Office.js is pre-loaded, no DOM)
// Edit the SIGNATURE_HTML block below to update the signature for all users.

const SIGNATURE_HTML = `
<table cellpadding="0" cellspacing="0" border="0"
  style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#333333;margin-top:12px;">
  <tr>
    <td style="padding-right:18px;vertical-align:top;padding-top:2px;">
      <img src="https://cyberitex.github.io/CTX-Signature/assets/logo.png"
           width="100" alt="CyberITEX" style="display:block;"/>
    </td>
    <td style="border-left:3px solid #0078D4;padding-left:16px;vertical-align:top;line-height:1.8;">
      <strong style="font-size:13pt;color:#222222;display:block;">Ahmed Ali</strong>
      <span style="color:#555555;display:block;">Security Engineer, <strong style="color:#222222;">CyberITEX</strong></span>
      <span style="color:#555555;display:block;margin-top:6px;">
        +1 970-460-8020 &nbsp;&nbsp;|&nbsp;&nbsp; +44 7-487-21-8887 &nbsp;&nbsp;|&nbsp;&nbsp;
        <a href="https://cyberitex.com" style="color:#0078D4;text-decoration:none;">cyberitex.com</a>
      </span>
      <span style="color:#555555;display:block;">
        <a href="mailto:Ahmed@cyberitex.com" style="color:#0078D4;text-decoration:none;">Ahmed@cyberitex.com</a>
        &nbsp;&nbsp;|&nbsp;&nbsp;
        <a href="https://cyberitex.com/appointments" style="color:#0078D4;text-decoration:none;">cyberitex.com/appointments</a>
      </span>
    </td>
  </tr>
</table>`;

function onNewMessageCompose(event) {
  var done = false;
  function complete() {
    if (!done) { done = true; event.completed({ allowEvent: true }); }
  }

  setTimeout(complete, 4000);

  Office.context.mailbox.item.body.setSignatureAsync(
    SIGNATURE_HTML,
    { coercionType: Office.CoercionType.Html },
    function () { complete(); }
  );
}

// Register function names so Outlook can invoke them from the manifest
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
