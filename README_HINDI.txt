सेवा हाजरी — GitHub Pages पैकेज

इस ZIP में:
1. index.html — वेबसाइट का 2-page जैसा interface
2. data.json — FINAL sheet से निकाला गया प्रारम्भिक data
3. update_data.py — OneDrive Excel की FINAL sheet से data.json अपडेट करता है
4. .github/workflows/update.yml — हर 10 मिनट में update check
5. assets/rs-sb-logo.png — RS/SB logo

FINAL sheet mapping:
B = Badge No.
C = Sewadar Name
D = Father/Husband Name
H = Day
I = Night
J = Saturday
K = Sunday
L = Lipai
M = Beas

जरूरी:
- OneDrive file का sharing “Anyone with the link can view” होना चाहिए, ताकि GitHub Action Excel डाउनलोड कर सके।
- GitHub repository में इन सभी files को root में रखें और .github/workflows/update.yml को उसी path पर रखें।
- GitHub Pages को main branch / root से deploy करें।
- Search box में typed text automatically UPPERCASE होगा।
- Last 5 searched IDs browser के localStorage में रहेंगे।
