# Fresher Party Event Manager

College Fresher & Farewell event ke liye live tracker portal — pure HTML + Tailwind CSS
(CDN), backend ke roop me Google Apps Script + Google Sheets.

## Pages

| File | Kaam |
|---|---|
| `index.html` | Dashboard — meetings, gallery, volunteers, feedback, about |
| `collection-entry.html` | Collection entry aur records |
| `expenses.html` | Expense entry, bills aur budget summary |
| `tasks.html` | Tasks / programs list |
| `participants.html` | Participant registration (list sirf admin ko) |
| `access.html` | Login, volunteer registration, admin panel |
| `Code.gs` | Apps Script backend |

## Setup (ek baar)

1. **Apps Script**: `Code.gs` ka poora content Apps Script editor me paste karein.
2. **Admin PIN**: Project Settings → Script Properties → property add karein:
   - Name: `ADMIN_PIN`
   - Value: apna PIN (jo bhi rakhna ho)

   > PIN ab kisi bhi HTML file me nahi hai. Use kabhi bhi HTML me wapas na daalein.
  > QR files ko kisi specific Drive folder me rakhna ho to `Code.gs` me `QR_FOLDER_ID` me us folder ka ID daalein. Dance ke MP3 songs ke liye `SONG_FOLDER_ID` set kar sakte hain. Khali chhodne par respective files Drive root folder me save hongi.
3. **Deploy**: Deploy → New deployment → Web app
   - Execute as: **Me**
   - Who has access: **Anyone**
4. **API URL**: nayi `/exec` URL copy karke har HTML file ke `CONFIG` block ki
   `API_URL` line me paste karein (6 files, har file me ek hi line).

## Security notes

Kya theek kiya gaya:

- Admin PIN server-side (Script Properties) me hai, page source me nahi.
- Login par server session token deta hai; har delete / approve / upload server par
  verify hota hai. Browser me `role` badalne se koi access nahi milta.
- Volunteer passwords `sha256$<salt>$<hash>` ke roop me store hote hain. Purane
  plaintext passwords pehle successful login par apne aap hash ho jate hain.
- Volunteer login POST se hota hai, isliye password URL / browser history me nahi jata.
- Sheet ka sara text HTML me daalne se pehle escape hota hai (stored XSS band).
- Buttons event delegation use karte hain, inline `onclick` nahi.
- Bill aur media links sirf `http`/`https` allow karte hain.
- Logout par server session bhi khatam hota hai, aur sirf is portal ki keys hatti hain.
- Volunteer registration me QR image Drive par save hoti hai; login ke baad volunteer ka apna QR collection page par dikhta hai.
- Har nayi volunteer collection entry me collector phone/name save hota hai, jisse Access page par volunteer-wise total aur entry count admin ko live milta hai.

Abhi baaki hai (jaan-bujh kar chhoda gaya):

- **`getData` public hai.** Jo bhi API URL jaanta hai wo bina login ke participants
  aur vendors ke phone numbers padh sakta hai. Isko band karne ke liye `ACTION_ROLES`
  me `getData: ["Admin", "Volunteer"]` add karna hoga.
- Drive par upload ki gayi files "anyone with the link" par share hoti hain, kyunki
  gallery aur bill links public page par dikhte hain. Personal documents upload na karein.

## Note

VS Code `@apply` par "Unknown at rule" warning dikha sakta hai — ye normal hai,
Tailwind CDN ise runtime par process kar deta hai.
