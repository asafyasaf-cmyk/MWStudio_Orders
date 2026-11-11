# שמירת הזמנות ב-Google Sheets

כל הזמנה מהאתר נוספת כשורה בגיליון, עם מספר הזמנה ועמודת סטטוס לבחירה:
**התקבלה · בטיפול · בייצור · מוכנה · נמסרה · בוטלה** (כל סטטוס בצבע אחר).

## התקנה (פעם אחת, כ-5 דקות)
1. פותחים גיליון חדש: https://sheets.new ונותנים לו שם, למשל "הזמנות MW Studio".
2. בתפריט: **הרחבות (Extensions) ← Apps Script**.
3. מוחקים את מה שכתוב שם ומדביקים את כל התוכן של [`Code.gs`](Code.gs). שומרים (Ctrl+S).
4. למעלה בוחרים את הפונקציה **setup** ולוחצים **▶ Run**. Google יבקש אישור הרשאות לגיליון: **Review permissions** ← בוחרים את החשבון ← **Advanced** ← **Go to … (unsafe)** ← **Allow**.
   (זו האזהרה הרגילה לסקריפט פרטי שלא עבר אימות של Google. הסקריפט כותב רק לגיליון הזה.)
5. **Deploy ← New deployment** ← בגלגל השיניים בוחרים **Web app**:
   - **Execute as:** Me
   - **Who has access:** Anyone
   - **Deploy**
6. מעתיקים את ה-**Web app URL** (מתחיל ב-`https://script.google.com/macros/s/`) ושולחים לי, או מדביקים בעצמכם ב-`index.html`:
   ```js
   const ORDERS_ENDPOINT='https://script.google.com/macros/s/.../exec';
   ```

## התראה במייל (לא חובה)
ב-`Code.gs` ממלאים `const NOTIFY_EMAIL='you@example.com';`, שומרים, ואז **Deploy ← Manage deployments ← ✏️ ← Version: New version ← Deploy** (הכתובת נשארת אותה כתובת).

## חשוב לדעת
- כל שינוי בסקריפט נכנס לתוקף רק אחרי **New version** ב-Manage deployments.
- הכתובת מקבלת הזמנות מכל אחד (כמו כל טופס באתר). יש מלכודת בוטים פשוטה, והסקריפט מונע נוסחאות זדוניות בתאים.
- אם השמירה נכשלת, הלקוח רואה הודעה ועדיין יכול לשלוח בוואטסאפ.
