THRYVE PULSE PAYMENT + CASH COUPON DEPLOYMENT

This complete package includes the Netlify Functions. The earlier package was missing these functions, which is why online payment and cash coupon redemption could not work.

GITHUB FIRST
1. Extract the ZIP.
2. Upload the CONTENTS of the Thryve_Pulse_Cash_Coupon_FIXED folder to the repository root, preserving the netlify/functions folder.
3. Commit and push. Do not test the payment flow on GitHub Pages; it cannot run Netlify Functions.
4. Connect the repository to Netlify and deploy.

NETLIFY ENVIRONMENT VARIABLES (Production context and Functions scope)
- RAZORPAY_KEY_ID
- RAZORPAY_KEY_SECRET (same Razorpay mode as the key ID)
- THRYVE_TOKEN_SECRET (keep stable and private)
- APPS_SCRIPT_URL (Apps Script Web App URL ending /exec)
After changing variables, redeploy.

GOOGLE APPS SCRIPT
1. Open the Apps Script project bound to the correct spreadsheet.
2. Replace its code with Thryve_Pulse_Cash_Coupon_Updated_Code.gs, save, then deploy a new Web App version. Keep access available to anyone if Netlify must call it anonymously.
3. Run testGmailAuth once if Gmail access hasn't been authorized.
4. Run generateCashCoupons from the Apps Script editor to create a batch of active unique coupons in the Cash Coupons tab. Give a code only after receiving cash.

TEST
- Test Razorpay using matching Test Mode credentials first.
- Successful payment should append a row to Registrations and email the new assessment link.
- An active TP-CASH-XXXXXXXXXXXX coupon should redeem once, mark its row Redeemed, append a registration row and send the link email. Reuse must fail.
- The assessment link domain is https://thryvemeeraki.github.io/thryvepulse2.o/thryve-assessment.html?token=...

SECURITY
Never put Razorpay secret or THRYVE_TOKEN_SECRET into HTML or public GitHub source. Set them only in Netlify environment variables.
