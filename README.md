# LeaveDesk

    npm install
    npm run dev        # http://localhost:5173

Before first use, in Firebase Console:
1. Authentication → Sign-in method → enable Email/Password
2. Firestore Database → create database
3. Firestore Database → Rules → paste the contents of `firestore.rules` → Publish

Deploy (Firebase Hosting + rules):

    npm run deploy     # runs `firebase login` first time: npx firebase-tools login
