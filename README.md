# Firma 2026 — admin panel

Haydovchilar, yo'l varaqalari, ichki xodimlar, firmalar, to'lovlar va fond uchun admin panel.
Interfeys uch alifboda ishlaydi: lotin, kirill va yangi lotin (Ö, Ğ, Ş, Ç).

## Fayllar
- `index.html` — saytning o'zi (GitHub Pages shu faylni ochadi)
- `support.js` — sahifa ishlashi uchun yordamchi skript
- `apps-script/Code.gs` — Google Sheets bazasi uchun Apps Script kodi

## Google Sheets'ni ulash
1. Google Sheets'da yangi jadval oching (masalan, «Firma 2026 baza»).
2. **Kengaytmalar → Apps Script**. Ichidagi kodni o'chirib, `apps-script/Code.gs` ni to'liq qo'ying va saqlang.
3. Xohlasangiz, koddagi `PAROL` qatoriga parol yozing.
4. **Joylashtirish → Yangi joylashtirish → Veb-ilova**: «Kim nomidan» — Men, «Kim kira oladi» — Hamma. Ruxsat bering.
5. Chiqqan veb-ilova URL'ini (…/exec) saytdagi **Sozlamalar** sahifasiga qo'ying, parolni yozing va **Saqlash va ulanish** ni bosing.
6. Birinchi marta **Hozirgi ma'lumotlarni Sheets'ga yuborish** tugmasini bosing.

Shundan keyin har bir o'zgarish avtomatik Google Sheets'ga yoziladi va boshqa kompyuterlarda ham ko'rinadi
(har bir kompyuterda Sozlamalar'ga bir marta URL va parolni kiritish kerak).

Apps Script kodi o'zgartirilsa: **Joylashtirish → Joylashtirishlarni boshqarish → tahrirlash → Yangi versiya**.
