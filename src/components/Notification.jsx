import React, { useEffect } from 'react';

const Notification = ({ message, type, onClose }) => {
  useEffect(() => {
    // message එකක් නැත්නම් ටයිමර් එක වැඩ කරන්න අවශ්‍ය නැහැ
    if (!message) return;

    const timer = setTimeout(() => {
      onClose();
    }, 3000);
    return () => clearTimeout(timer);
  }, [message, onClose]);

  // message එකක් නැත්නම් කිසිවක් return කරන්න එපා
  if (!message) return null;

  const bgColor = type === 'success' ? 'bg-green-500' : 'bg-red-500';

  return (
    <div className={`fixed top-5 right-5 ${bgColor} text-white px-6 py-3 rounded-lg shadow-lg z-50 animate-fade-in-down`}>
      {message}
    </div>
  );
};

export default Notification;