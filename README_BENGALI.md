# Excel Free AI Assistant — বাংলা সেটআপ

## কী লাগবে
- Excel 2024 LTSC (Windows)
- ইন্টারনেট
- Gemini API key (Google AI Studio-তে API access ও free-tier availability আপনার account/model/region অনুযায়ী যাচাই করুন)
- একটি GitHub account (add-in-এর web files HTTPS-এ host করতে)

## গুরুত্বপূর্ণ
এটি Microsoft 365 Copilot নয়। এটি একটি personal Office task-pane add-in, যা Gemini API ব্যবহার করে। API free tier-এ rate/usage limits থাকতে পারে। API key কাউকে দেবেন না এবং sensitive worksheet data AI-তে পাঠাবেন না।

## ধাপ ১ — API key
Google AI Studio থেকে API key তৈরি করুন। API key কোনো public GitHub repository-তে কখনো লিখবেন না।

## ধাপ ২ — GitHub Pages-এ web files host
1. GitHub-এ নতুন repository বানান: `excel-ai-assistant`
2. `taskpane.html`, `taskpane.css`, `taskpane.js` ফাইলগুলো repository-তে upload করুন।
3. Repository Settings → Pages → Deploy from a branch → `main` এবং `/root` নির্বাচন → Save.
4. Pages-এর HTTPS URL তৈরি হওয়া পর্যন্ত অপেক্ষা করুন; যেমন `https://USERNAME.github.io/excel-ai-assistant/`

## ধাপ ৩ — Manifest URL ঠিক করুন
`manifest.xml` খুলুন। `YOUR-GITHUB-USERNAME`-এর সব জায়গায় নিজের GitHub username বসান। URL path-এ repository নাম ঠিক রাখুন। Save করুন।

## ধাপ ৪ — Excel-এ add-in sideload
Excel খুলুন → Insert → Add-ins / My Add-ins → Manage My Add-ins → Upload My Add-in (নাম/version অনুযায়ী menu কিছুটা আলাদা হতে পারে) → `manifest.xml` নির্বাচন → Upload.
যদি Upload My Add-in অপশন না থাকে, Office-এর add-in sideload/developer menu policy দ্বারা বন্ধ থাকতে পারে। Excel build এবং organization policy পরীক্ষা করুন।

## ব্যবহার
- Excel ribbon থেকে add-in খুলুন।
- Gemini API key paste করুন।
- Worksheet-এ data range select করুন।
- Read selected cells চাপুন।
- প্রশ্ন লিখুন এবং Send to AI চাপুন।
- AI উত্তর cell-এ বসাতে চাইলে target cell select করে Insert answer in cell চাপুন।

## গোপনীয়তা ও নিরাপত্তা
- Key browser session-এ রাখা হয়, স্থায়ীভাবে save করা হয় না।
- Data ও prompt সরাসরি Gemini API-তে পাঠানো হয়; Google-এর API terms/privacy policy প্রযোজ্য।
- Public hosting-এর code-এ আপনার key নেই। তবে key যে browser-এ ব্যবহার করছেন সেটি নিরাপদ রাখুন।
- AI-এর formula/code চালানোর আগে যাচাই করুন।
- এই prototype-এর API call client-side; personal testing-এর জন্য। Public বা multi-user deployment-এ server-side proxy ও authentication ব্যবহার করুন।

## Troubleshooting
- Add-in load না হলে HTTPS Pages URL সঠিক কিনা দেখুন এবং Excel বন্ধ করে আবার খুলুন।
- API error 400: model/API endpoint বা request check করুন।
- 403: key permission, region, API enablement, restrictions check করুন।
- 429: quota/rate limit; পরে চেষ্টা করুন বা provider quota দেখুন।
- Excel API error: Excel update/build এবং Office.js support পরীক্ষা করুন।
