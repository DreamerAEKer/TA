// Firebase Configuration
// TODO: Replace with your actual Firebase project credentials
// 1. Go to Firebase Console (https://console.firebase.google.com/)
// 2. Create a new project or select an existing one.
// 3. Add a Web App to the project and copy the firebaseConfig object below.
// 4. Enable Firestore Database in your Firebase project.

const firebaseConfig = {
    apiKey: "YOUR_API_KEY",
    authDomain: "YOUR_PROJECT_ID.firebaseapp.com",
    projectId: "YOUR_PROJECT_ID",
    storageBucket: "YOUR_PROJECT_ID.firebasestorage.app",
    messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
    appId: "YOUR_APP_ID"
};

// Initialize Firebase
if (typeof firebase !== "undefined") {
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    window.db = db;
} else {
    console.warn("Firebase SDK not loaded");
}

