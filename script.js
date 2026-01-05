// UI Interactions
$("#lengthSlider").on("input", function () {
  const labels = ["Short", "Normal", "Detailed"];
  $("#lengthLabel").text(labels[$(this).val() - 1]);
});

// Button Bindings
$("#replyBtn").click(() => runTask("reply"));
$("#rewriteBtn").click(() => runTask("rewrite"));
$("#arabicBtn").click(() => runTask("arabic"));
$("#summarizeBtn").click(() => runTask("summarize"));
$("#shortenBtn").click(() => runTask("shorten"));
$("#customBtn").click(() => runTask("custom", $("#customPrompt").val()));

// Auto-save settings
$("#apiKeyInput, #styleInput").on("input", function () {
  localStorage.setItem("geminiKey", $("#apiKeyInput").val());
  localStorage.setItem("myStyle", $("#styleInput").val());
});

Office.onReady(() => {
  $("#apiKeyInput").val(localStorage.getItem("geminiKey"));
  $("#styleInput").val(localStorage.getItem("myStyle"));
});

async function runTask(type, customText = "") {
  const key = $("#apiKeyInput").val();
  if (!key) return updateStatus("🔑 Enter Key!", true);

  const style = $("#styleInput").val() || "Professional";
  const tone = $("#toneSelect").val();
  const len = $("#lengthLabel").text();

  let prompt = `Persona: ${style}. Tone: ${tone}. Length: ${len}. `;
  if (type === "reply") prompt += "Draft a response to this email history.";
  else if (type === "rewrite") prompt += "Rewrite this text for better flow.";
  else if (type === "arabic") prompt = "Translate this to professional Arabic.";
  else if (type === "summarize") prompt = "List key points in bullets.";
  else if (type === "shorten") prompt = "Make this very short.";
  else prompt = customText;

  updateStatus("⌛ AI Thinking...");

  Office.context.mailbox.item.body.getAsync("text", async (res) => {
    try {
      const result = await callGemini(key, res.value, prompt);
      Office.context.mailbox.item.body.setSelectedDataAsync(result, { coercionType: "text" }, (asyncResult) => {
        if (asyncResult.status === "failed") {
          Office.context.mailbox.item.body.prependAsync(result + "\n\n---\n");
        }
        updateStatus("✅ Success!");
      });
    } catch (e) {
      updateStatus("❌ Error", true);
    }
  });
}

async function callGemini(key, text, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${key}`;
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ contents: [{ parts: [{ text: `Task: ${prompt}\n\nContent: ${text}` }] }] }),
  });
  const data = await response.json();
  return data.candidates[0].content.parts[0].text;
}

function updateStatus(msg, isError = false) {
  $("#status")
    .text(msg)
    .css("background", isError ? "#fde7e9" : "#dff6dd")
    .css("color", isError ? "#a4262c" : "#107c10");
}
