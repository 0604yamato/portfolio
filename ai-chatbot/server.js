
/**
 * Minimal Dify proxy (blocking mode) with LINE Bot integration
 * Usage:
 *   1) npm init -y && npm i express cors node-fetch dotenv @line/bot-sdk
 *   2) create .env with DIFY_API_KEY=xxxxx, LINE_CHANNEL_SECRET=xxxxx, LINE_CHANNEL_ACCESS_TOKEN=xxxxx
 *   3) node server.js
 *   4) Open the HTML file and set ENDPOINT to 'http://localhost:8787/api/chat'
 */
import express from "express";
import cors from "cors";
import fetch from "node-fetch";
import dotenv from "dotenv";
import { messagingApi, middleware } from "@line/bot-sdk";
dotenv.config();

const app = express();
app.use(cors());               // allow local file origin
app.use(express.json());

// Serve static files from 'public' directory
app.use(express.static('public'));

const PORT = process.env.PORT || 8787;
const DIFY_API_KEY = process.env.DIFY_API_KEY;
const DIFY_URL = process.env.DIFY_URL || "https://api.dify.ai/v1/chat-messages";

// LINE Bot configuration
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;

if (!DIFY_API_KEY) {
  console.warn("WARNING: Missing DIFY_API_KEY in environment.");
}

if (!LINE_CHANNEL_SECRET || !LINE_CHANNEL_ACCESS_TOKEN) {
  console.warn("WARNING: Missing LINE credentials in environment.");
}

// Initialize LINE client
const lineConfig = {
  channelAccessToken: LINE_CHANNEL_ACCESS_TOKEN || "",
  channelSecret: LINE_CHANNEL_SECRET || ""
};
const lineClient = new messagingApi.MessagingApiClient({
  channelAccessToken: lineConfig.channelAccessToken
});

app.post("/api/chat", async (req, res) => {
  try {
    const { query, conversation_id, user_id, inputs } = req.body || {};
    if (!query || typeof query !== "string") {
      return res.status(400).json({ error: "query (string) is required" });
    }
    const payload = {
      query,
      inputs: inputs || {},
      response_mode: "streaming",   // use streaming mode for Agent Chat
      user: user_id || "local-user",
      conversation_id: conversation_id || null,
    };
    const r = await fetch(DIFY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${DIFY_API_KEY}`
      },
      body: JSON.stringify(payload)
    });

    if (!r.ok) {
      const errorData = await r.json();
      console.error("Dify error:", errorData);
      return res.status(500).json({ error: "Dify API error", detail: errorData });
    }

    // Handle streaming response
    let fullAnswer = "";
    let conversationId = null;
    const text = await r.text();

    // Parse SSE format (data: {...}\n\n)
    const lines = text.split('\n');
    for (const line of lines) {
      if (line.startsWith('data: ')) {
        try {
          const jsonStr = line.slice(6); // Remove 'data: ' prefix
          const data = JSON.parse(jsonStr);

          if (data.event === 'agent_message' || data.event === 'message') {
            fullAnswer += data.answer || "";
          }
          if (data.conversation_id) {
            conversationId = data.conversation_id;
          }
        } catch (e) {
          // Skip invalid JSON lines
        }
      }
    }

    console.log("=== Dify Response Debug ===");
    console.log("Answer:", fullAnswer);
    console.log("Answer length:", fullAnswer.length);
    console.log("Conversation ID:", conversationId);
    console.log("===========================");

    return res.json({
      answer: fullAnswer || "",
      conversation_id: conversationId || null
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Server error" });
  }
});

// LINE Webhook endpoint
app.post("/webhook", async (req, res) => {
  try {
    // Verify LINE signature
    const signature = req.get("x-line-signature");
    if (!signature) {
      console.error("Missing LINE signature");
      return res.status(401).send("Unauthorized");
    }

    // Validate webhook signature
    const body = JSON.stringify(req.body);
    const crypto = await import("crypto");
    const hash = crypto.createHmac("sha256", LINE_CHANNEL_SECRET)
      .update(body)
      .digest("base64");

    if (hash !== signature) {
      console.error("Invalid LINE signature");
      return res.status(401).send("Unauthorized");
    }

    const events = req.body.events;
    console.log("=== LINE Webhook Events ===");
    console.log(JSON.stringify(events, null, 2));

    // Process each event
    await Promise.all(events.map(async (event) => {
      if (event.type !== "message" || event.message.type !== "text") {
        return;
      }

      const userMessage = event.message.text;
      const userId = event.source.userId;
      const replyToken = event.replyToken;

      console.log(`User ${userId}: ${userMessage}`);

      try {
        // Call Dify API
        const difyPayload = {
          query: userMessage,
          inputs: {},
          response_mode: "streaming",
          user: userId,
          conversation_id: null
        };

        const difyResponse = await fetch(DIFY_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${DIFY_API_KEY}`
          },
          body: JSON.stringify(difyPayload)
        });

        if (!difyResponse.ok) {
          const errorData = await difyResponse.json();
          console.error("Dify error:", errorData);
          throw new Error("Dify API error");
        }

        // Handle streaming response
        let botReply = "";
        const text = await difyResponse.text();
        const lines = text.split('\n');
        for (const line of lines) {
          if (line.startsWith('data: ')) {
            try {
              const jsonStr = line.slice(6);
              const data = JSON.parse(jsonStr);
              if (data.event === 'agent_message' || data.event === 'message') {
                botReply += data.answer || "";
              }
            } catch (e) {
              // Skip invalid JSON
            }
          }
        }

        if (!botReply) {
          botReply = "申し訳ございません。応答を生成できませんでした。";
        }

        console.log(`Bot reply: ${botReply}`);

        // Reply to LINE
        await lineClient.replyMessage({
          replyToken: replyToken,
          messages: [{
            type: "text",
            text: botReply
          }]
        });

      } catch (error) {
        console.error("Error processing message:", error);
        // Try to send error message to user
        try {
          await lineClient.replyMessage({
            replyToken: replyToken,
            messages: [{
              type: "text",
              text: "申し訳ございません。エラーが発生しました。"
            }]
          });
        } catch (replyError) {
          console.error("Error sending error message:", replyError);
        }
      }
    }));

    res.status(200).send("OK");
  } catch (err) {
    console.error("Webhook error:", err);
    res.status(500).send("Server error");
  }
});

app.listen(PORT, () => {
  console.log(`Proxy listening on http://localhost:${PORT}`);
  console.log(`Webhook endpoint: http://localhost:${PORT}/webhook`);
});
