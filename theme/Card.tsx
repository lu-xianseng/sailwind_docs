import React from 'react';

interface CardProps {
  title?: React.ReactNode;
  icon?: React.ReactNode;
  content?: React.ReactNode;
  style?: React.CSSProperties;
}

export function Card({ title, icon, content, style }: CardProps) {
  return (
    <div
      style={{
        padding: '16px',
        display: 'flex',
        flexDirection: 'column',
        gap: '12px',
        ...style,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
        {icon}
        <div style={{ fontSize: '18px', fontWeight: 600 }}>{title}</div>
      </div>
      <div>{content}</div>
    </div>
  );
}