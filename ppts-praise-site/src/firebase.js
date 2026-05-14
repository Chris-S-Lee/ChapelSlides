// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyBFFReEIOjMG2NOH4QQZjjCCyVtx5MWUaE",
  authDomain: "chapelslides-eb5c0.firebaseapp.com",
  projectId: "chapelslides-eb5c0",
  storageBucket: "chapelslides-eb5c0.firebasestorage.app",
  messagingSenderId: "711394611559",
  appId: "1:711394611559:web:176f07af395953ba2e1fc2"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);