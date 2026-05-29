// Gmail API via Google Apps Script

const SECRET = "daxos-gmail-2026";
const SCRIPT_ID = "17InArrCTsC3SM2-R8EiEFkv2Sl5QJruSnNIi5jFe36iaXLo_eAS4Tdgh";

// Exact signature HTML from a real sent email — never passes through URL encoding
const SIGNATURE_HTML = '<div><div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature"><div dir="ltr"><span style="font-size:x-small;color:rgb(0,0,0)"><font face="monospace"><p style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline">Mark Davidoff</span></p><p style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline">Chief Investment Officer</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline">Email: cus@daxos.capital</span><span style="vertical-align:baseline"> </span></p><p style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline">Telegram: @mark_daxos</span></p></font></span><span style="font-size:x-small;color:rgb(0,0,0)"><font face="monospace"><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline"><br></span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="vertical-align:baseline">www.daxos.capital</span></p></font></span><span style="font-size:x-small;color:rgb(0,0,0)"></span><span style="font-size:x-small;color:rgb(0,0,0)"><font face="monospace"><span style="vertical-align:baseline">NYC|MIA</span></font></span></div></div></div>';

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  e.parameter = body;
  return doGet(e);
}

function doGet(e) {
  var p = e.parameter;

  if (p.key !== SECRET) {
    return json({ error: "unauthorized" });
  }

  var action = p.action;

  try {
    if (action === "search") {
      return json(searchEmails(p.q, parseInt(p.max) || 10));
    }
    if (action === "read") {
      return json(readEmail(p.id));
    }
    if (action === "send") {
      return json(sendEmail(p.to, p.subject, p.body, p.cc, p.bcc, p.attachments));
    }
    if (action === "draft") {
      return json(createDraft(p.to, p.subject, p.body, p.cc, p.attachments));
    }
    if (action === "draft_outreach") {
      return json(draftOutreach(p.to, p.subject, p.body, p.cc, p.bcc, p.attachments));
    }
    if (action === "reply") {
      return json(replyToThread(p.threadId, p.body, p.attachments));
    }
    if (action === "draft_reply") {
      return json(draftReply(p.threadId, p.body, p.attachments));
    }
    if (action === "star") {
      return json(starMessage(p.id));
    }
    if (action === "unstar") {
      return json(unstarMessage(p.id));
    }
    if (action === "mark_read") {
      return json(markRead(p.id));
    }
    if (action === "trash") {
      return json(trashMessage(p.id));
    }
    if (action === "create_label") {
      return json(createLabel(p.name));
    }
    if (action === "add_label") {
      return json(addLabel(p.threadId, p.name));
    }
    if (action === "remove_label") {
      return json(removeLabel(p.threadId, p.name));
    }
    if (action === "threads") {
      return json(getThreads(p.label || "INBOX", parseInt(p.max) || 10));
    }
    if (action === "labels") {
      return json(getLabels());
    }
    if (action === "get_attachment") {
      return json(getAttachment(p.id, p.filename));
    }
    if (action === "update_code") {
      return json(updateCode(p.url));
    }
    return json({ error: "unknown action" });
  } catch (err) {
    return json({ error: err.toString() });
  }
}

function attachmentsToBlobs(attachments) {
  if (!attachments || !attachments.length) return null;
  return attachments.map(function(a) {
    var bytes = Utilities.base64Decode(a.content);
    var mime = a.mimeType || 'application/octet-stream';
    var name = a.filename || 'attachment';
    return Utilities.newBlob(bytes, mime, name);
  });
}

function draftOutreach(to, subject, body, cc, bcc, attachments) {
  var htmlBody = '<div dir="ltr"><div>' + body.replace(/\n/g, '<br>') + '<br></div>' + SIGNATURE_HTML + '</div>';
  var options = { htmlBody: htmlBody };
  if (cc) options.cc = cc;
  if (bcc) options.bcc = bcc;
  var blobs = attachmentsToBlobs(attachments);
  if (blobs) options.attachments = blobs;
  var draft = GmailApp.createDraft(to, subject, body, options);
  return { success: true, draftId: draft.getId(), to: to, subject: subject, attachmentCount: blobs ? blobs.length : 0 };
}

function createDraft(to, subject, body, cc, attachments) {
  var htmlBody = '<div dir="ltr"><div>' + body.replace(/\n/g, '<br>') + '<br></div>' + SIGNATURE_HTML + '</div>';
  var options = { htmlBody: htmlBody };
  if (cc) options.cc = cc;
  var blobs = attachmentsToBlobs(attachments);
  if (blobs) options.attachments = blobs;
  var draft = GmailApp.createDraft(to, subject, body, options);
  return { success: true, draftId: draft.getId(), to: to, subject: subject, attachmentCount: blobs ? blobs.length : 0 };
}

function searchEmails(query, max) {
  var threads = GmailApp.search(query, 0, max);
  var results = [];
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var last = msgs[msgs.length - 1];
    results.push({
      threadId: threads[i].getId(),
      messageId: last.getId(),
      subject: last.getSubject(),
      from: last.getFrom(),
      to: last.getTo(),
      date: last.getDate().toISOString(),
      snippet: last.getPlainBody().substring(0, 200),
      unread: last.isUnread(),
      messageCount: msgs.length
    });
  }
  return { results: results, count: results.length, query: query };
}

function readEmail(id) {
  var msg = GmailApp.getMessageById(id);
  if (!msg) return { error: "message not found" };
  return {
    id: msg.getId(),
    threadId: msg.getThread().getId(),
    subject: msg.getSubject(),
    from: msg.getFrom(),
    to: msg.getTo(),
    cc: msg.getCc(),
    date: msg.getDate().toISOString(),
    body: msg.getPlainBody(),
    htmlBody: msg.getBody().substring(0, 5000),
    labels: msg.getThread().getLabels().map(function(l) { return l.getName(); }),
    attachments: msg.getAttachments().map(function(a) { return { name: a.getName(), size: a.getSize() }; })
  };
}

function sendEmail(to, subject, body, cc, bcc, attachments) {
  var htmlBody = '<div dir="ltr"><div>' + body.replace(/\n/g, '<br>') + '<br></div>' + SIGNATURE_HTML + '</div>';
  var options = { htmlBody: htmlBody };
  if (cc) options.cc = cc;
  if (bcc) options.bcc = bcc;
  var blobs = attachmentsToBlobs(attachments);
  if (blobs) options.attachments = blobs;
  GmailApp.sendEmail(to, subject, body, options);
  return { success: true, to: to, subject: subject, attachmentCount: blobs ? blobs.length : 0 };
}

function replyToThread(threadId, body, attachments) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: "thread not found" };
  var last = thread.getMessages()[thread.getMessageCount() - 1];
  var options = {};
  var blobs = attachmentsToBlobs(attachments);
  if (blobs) options.attachments = blobs;
  last.reply(body, options);
  return { success: true, threadId: threadId, subject: thread.getFirstMessageSubject(), attachmentCount: blobs ? blobs.length : 0 };
}

function draftReply(threadId, body, attachments) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: "thread not found" };
  var last = thread.getMessages()[thread.getMessageCount() - 1];
  var options = {};
  var blobs = attachmentsToBlobs(attachments);
  if (blobs) options.attachments = blobs;
  var draft = last.createDraftReply(body, options);
  return { success: true, draftId: draft.getId(), threadId: threadId, subject: thread.getFirstMessageSubject(), attachmentCount: blobs ? blobs.length : 0 };
}

function starMessage(id) {
  var msg = GmailApp.getMessageById(id);
  if (!msg) return { error: "message not found" };
  msg.star();
  return { success: true, id: id };
}

function unstarMessage(id) {
  var msg = GmailApp.getMessageById(id);
  if (!msg) return { error: "message not found" };
  msg.unstar();
  return { success: true, id: id };
}

function markRead(id) {
  var msg = GmailApp.getMessageById(id);
  if (!msg) return { error: "message not found" };
  msg.markRead();
  return { success: true, id: id };
}

function trashMessage(id) {
  var msg = GmailApp.getMessageById(id);
  if (!msg) return { error: "message not found" };
  msg.moveToTrash();
  return { success: true, id: id };
}

function createLabel(name) {
  var label = GmailApp.createLabel(name);
  return { success: true, name: label.getName() };
}

function addLabel(threadId, name) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: "thread not found" };
  var label = GmailApp.getUserLabelByName(name);
  if (!label) return { error: "label not found: " + name };
  thread.addLabel(label);
  return { success: true, threadId: threadId, label: name };
}

function removeLabel(threadId, name) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: "thread not found" };
  var label = GmailApp.getUserLabelByName(name);
  if (!label) return { error: "label not found: " + name };
  thread.removeLabel(label);
  return { success: true, threadId: threadId, label: name };
}

function getThreads(label, max) {
  var threads;
  if (label === "INBOX") {
    threads = GmailApp.getInboxThreads(0, max);
  } else if (label === "STARRED") {
    threads = GmailApp.getStarredThreads(0, max);
  } else if (label === "UNREAD") {
    threads = GmailApp.search("is:unread", 0, max);
  } else {
    threads = GmailApp.search("label:" + label, 0, max);
  }

  var results = [];
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var last = msgs[msgs.length - 1];
    results.push({
      threadId: threads[i].getId(),
      messageId: last.getId(),
      subject: last.getSubject(),
      from: last.getFrom(),
      date: last.getDate().toISOString(),
      snippet: last.getPlainBody().substring(0, 150),
      unread: last.isUnread(),
      messageCount: msgs.length
    });
  }
  return { label: label, results: results, count: results.length };
}

function getLabels() {
  var labels = GmailApp.getUserLabels();
  return {
    labels: labels.map(function(l) {
      return { name: l.getName(), unread: l.getUnreadCount() };
    })
  };
}

function getAttachment(messageId, filename) {
  var msg = GmailApp.getMessageById(messageId);
  if (!msg) return { error: "message not found" };
  var atts = msg.getAttachments();
  for (var i = 0; i < atts.length; i++) {
    if (!filename || atts[i].getName() === filename) {
      var bytes = atts[i].getBytes();
      return {
        success: true,
        name: atts[i].getName(),
        contentType: atts[i].getContentType(),
        size: bytes.length,
        base64: Utilities.base64Encode(bytes)
      };
    }
  }
  return { error: "attachment not found", available: atts.map(function(a){return a.getName();}) };
}

function updateCode(sourceUrl) {
  if (!sourceUrl) return { error: "url parameter required" };
  var token = ScriptApp.getOAuthToken();
  var baseApi = "https://script.googleapis.com/v1/projects/" + SCRIPT_ID;

  var newCode = UrlFetchApp.fetch(sourceUrl).getContentText();

  var currentResp = UrlFetchApp.fetch(baseApi + "/content", {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  var content = JSON.parse(currentResp.getContentText());

  var files = content.files;
  for (var i = 0; i < files.length; i++) {
    if (files[i].name === "Code" && files[i].type === "SERVER_JS") {
      files[i].source = newCode;
    }
  }

  var updateResp = UrlFetchApp.fetch(baseApi + "/content", {
    method: "put",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify({ files: files }),
    muteHttpExceptions: true
  });
  var updateResult = JSON.parse(updateResp.getContentText());
  if (updateResult.error) return { error: "update failed", details: updateResult.error };

  var versionResp = UrlFetchApp.fetch(baseApi + "/versions", {
    method: "post",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    payload: JSON.stringify({ description: "Auto update from Claude" }),
    muteHttpExceptions: true
  });
  var versionData = JSON.parse(versionResp.getContentText());
  if (!versionData.versionNumber) return { error: "version creation failed", details: versionData };

  var deploymentsResp = UrlFetchApp.fetch(baseApi + "/deployments", {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  var deployData = JSON.parse(deploymentsResp.getContentText());

  var deploymentId = null;
  for (var i = 0; i < deployData.deployments.length; i++) {
    var d = deployData.deployments[i];
    if (d.entryPoints) {
      for (var j = 0; j < d.entryPoints.length; j++) {
        if (d.entryPoints[j].entryPointType === "WEB_APP") {
          deploymentId = d.deploymentId;
        }
      }
    }
  }

  if (deploymentId) {
    var deployResp = UrlFetchApp.fetch(baseApi + "/deployments/" + deploymentId, {
      method: "put",
      headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
      payload: JSON.stringify({
        deploymentConfig: {
          scriptId: SCRIPT_ID,
          versionNumber: versionData.versionNumber,
          description: "v" + versionData.versionNumber
        }
      }),
      muteHttpExceptions: true
    });
    return {
      success: true,
      version: versionData.versionNumber,
      deploymentId: deploymentId,
      deployResult: JSON.parse(deployResp.getContentText())
    };
  }

  return { success: true, version: versionData.versionNumber, note: "version created but no web app deployment found to update" };
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
