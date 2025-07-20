// app/taskpane/page.tsx
"use client";

import { useEffect, useState } from "react";

export default function Taskpane() {
    const [emailBody, setEmailBody] = useState("");
    const [aiResponse, setAiResponse] = useState("");

    // Load Office.js runtime
    useEffect(() => {
        const script = document.createElement("script");
        script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
        script.async = true;
        script.onload = () => {
            if (window.Office) {
                Office.onReady((info) => {
                    if (info.host === Office.HostType.Outlook) {
                        const item = Office.context.mailbox?.item;
                        if (!item) {
                            console.warn("This is not a compose window. Cannot access email body.");
                            return;
                        }

                        item.body.getAsync("text", (result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                setEmailBody(result.value);
                            } else {
                                console.error("Failed to get body:", result.error);
                            }
                        });
                    }
                });
            }
        };
        document.body.appendChild(script);
    }, []);

    // Send to LLM API route
    async function generateReply() {
        const res = await fetch("/api/generate", {
            method: "POST",
            body: JSON.stringify({ prompt: emailBody }),
        });
        const { result } = await res.json();
        setAiResponse(result);
    }

    // Insert AI response into email
    function insertIntoEmail() {
        Office.context.mailbox.item.body.setAsync(aiResponse, {
            coercionType: Office.CoercionType.Html,
        });
    }

    return (
        <div className="p-4 space-y-4">
            <h1 className="text-xl font-bold">AI Email Assistant</h1>
            <textarea
                className="w-full h-32 p-2 border"
                value={emailBody}
                onChange={(e) => setEmailBody(e.target.value)}
            />
            <button onClick={generateReply} className="bg-blue-600 text-white px-4 py-2">
                Generate AI Response
            </button>
            {aiResponse && (
                <>
                    <h2 className="font-semibold">AI Suggestion:</h2>
                    <div className="p-2 border bg-gray-100">{aiResponse}</div>
                    <button
                        onClick={insertIntoEmail}
                        className="bg-green-600 text-white px-4 py-2 mt-2"
                    >
                        Insert into Email
                    </button>
                </>
            )}
        </div>
    );
}
