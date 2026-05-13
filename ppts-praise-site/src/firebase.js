// Import the functions you need from the SDKs you need
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyAwFH0u1Ccs-QpFYFBx_rbNs5dHLQxMMXI",
  authDomain: "chapelslides-4127f.firebaseapp.com",
  projectId: "chapelslides-4127f",
  storageBucket: "chapelslides-4127f.firebasestorage.app",
  messagingSenderId: "715626404737",
  appId: "1:715626404737:web:ce63459489760980f22680"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);