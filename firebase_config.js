// Firebase Configuration for Standalone App
// Replace with your actual Firebase project credentials
// Requires: Authentication (Optional), Firestore Database, Storage

const firebaseConfig = {
    apiKey: "AIzaSyDoq31ge5vIJqGmp1Ppq_rt1fZKMWVlRsc",
    authDomain: "tracking-analyst.firebaseapp.com",
    projectId: "tracking-analyst",
    storageBucket: "tracking-analyst.firebasestorage.app",
    messagingSenderId: "591722827387",
    appId: "1:591722827387:web:bae0979254d6ec751e5cde",
    measurementId: "G-9QV4SPDL9G"
};

// Initialize Firebase
if (typeof firebase !== "undefined") {
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    const storage = firebase.storage();
    window.db = db;
    window.storage = storage;
} else {
    console.error("Firebase SDK not loaded");
}

