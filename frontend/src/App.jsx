import React from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import LoginPage from './pages/LoginPage';
import CertificateGenerator from './pages/CertificateGenerator';
import { AuthProvider, useAuth } from './utils/auth';
import TESDAPage from './pages/TESDAPage';
import HomePage from './pages/HomePage';

function PrivateRoute({ children }) {
  const { isAuthenticated } = useAuth();
  return isAuthenticated ? children : <Navigate to="/login" replace />;
}

const App = () => (
  <AuthProvider>
    <Router>
      <Routes>
        <Route path="/" element={<Navigate to="/login" replace />} />
        <Route path="/login" element={<LoginPage />} />
        <Route path="/home" element={<PrivateRoute><HomePage /></PrivateRoute>} />
        <Route path="/generate" element={<PrivateRoute><CertificateGenerator /></PrivateRoute>} />
        <Route path="/tesda" element={<PrivateRoute><TESDAPage /></PrivateRoute>} />
      </Routes>
    </Router>
  </AuthProvider>
);

export default App;
