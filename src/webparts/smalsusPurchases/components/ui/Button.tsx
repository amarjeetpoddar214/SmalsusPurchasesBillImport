
import React from 'react';

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'ghost';
  size?: 'normal' | 'icon';
}

const Button: React.FC<ButtonProps> = ({
  children,
  className,
  variant = 'primary',
  size = 'normal',
  ...props
}) => {
  const baseClasses = 'inline-flex items-center justify-center rounded-md font-semibold transition-colors focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 focus:ring-offset-slate-900 disabled:opacity-50 disabled:pointer-events-none';
  
  const variantClasses = {
    primary: 'bg-indigo-600 text-white hover:bg-indigo-500',
    secondary: 'bg-slate-700 text-slate-100 hover:bg-slate-600',
    ghost: 'hover:bg-slate-700 text-slate-300 hover:text-white',
  };

  const sizeClasses = {
      normal: 'px-4 py-2 text-sm',
      icon: 'h-9 w-9',
  };

  return (
    <button
      className={`${baseClasses} ${variantClasses[variant]} ${sizeClasses[size]} ${className}`}
      {...props}
    >
      {children}
    </button>
  );
};

export default Button;
