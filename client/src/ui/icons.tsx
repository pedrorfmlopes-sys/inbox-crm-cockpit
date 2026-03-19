import React from "react";

/**
 * Modern SVG Icons for Universal Cockpit 3.0
 * Inspired by Lucide icons, styled for premium tech aesthetic.
 */

interface IconProps extends React.SVGProps<SVGSVGElement> {
    size?: number;
    color?: string;
    style?: React.CSSProperties;
    className?: string;
    title?: string;
}

const BaseIcon: React.FC<IconProps & { children: React.ReactNode }> = ({
    size = 20,
    color = "currentColor",
    strokeWidth = 2,
    children,
    ...props
}) => (
    <svg
        xmlns="http://www.w3.org/2000/svg"
        width={size}
        height={size}
        viewBox="0 0 24 24"
        fill="none"
        stroke={color}
        strokeWidth={strokeWidth}
        strokeLinecap="round"
        strokeLinejoin="round"
        {...props}
    >
        {children}
    </svg>
);

export const Sparkles: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z" />
        <path d="M5 3v4" />
        <path d="M19 17v4" />
        <path d="M3 5h4" />
        <path d="M17 19h4" />
    </BaseIcon>
);

export const Database: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <ellipse cx="12" cy="5" rx="9" ry="3" />
        <path d="M3 5V19A9 3 0 0 0 21 19V5" />
        <path d="M3 12A9 3 0 0 0 21 12" />
    </BaseIcon>
);

export const Files: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M15.5 2H8.6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h10.8c1.1 0 2-.9 2-2V7.5L15.5 2z" />
        <path d="M15 2v5h5" />
        <path d="M2 9v13c0 1.1.9 2 2 2h10" />
    </BaseIcon>
);

export const Settings: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z" />
        <circle cx="12" cy="12" r="3" />
    </BaseIcon>
);

export const Clipboard: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <rect width="8" height="4" x="8" y="2" rx="1" ry="1" />
        <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
    </BaseIcon>
);

export const ExternalLink: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
        <polyline points="15 3 21 3 21 9" />
        <line x1="10" x2="21" y1="14" y2="3" />
    </BaseIcon>
);

export const Edit: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
    </BaseIcon>
);

export const Receipt: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1-2-1Z" />
        <path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8" />
        <path d="M12 17.5V6.5" />
    </BaseIcon>
);

export const Handshake: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="m11 17 2 2 6-6" />
        <path d="m18 14 2.5 2.5a3.3 3.3 0 0 0 4.7-4.7" />
        <path d="M18 8c0 1.5-1.5 3-3.003 3.003h-2.197L8 15" />
        <path d="m5 10-4 4 3 3 4-4" />
        <path d="m4 6 5 5" />
        <path d="M8 5C8 3.5 9.5 2 11 2" />
        <path d="M15 13c1.5 0 3-1.5 3-3V5a3 3 0 0 0-6 0v2" />
    </BaseIcon>
);

export const Building: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <rect width="16" height="20" x="4" y="2" rx="2" ry="2" />
        <path d="M9 22v-4h6v4" />
        <path d="M8 6h.01" />
        <path d="M16 6h.01" />
        <path d="M8 10h.01" />
        <path d="M16 10h.01" />
        <path d="M8 14h.01" />
        <path d="M16 14h.01" />
    </BaseIcon>
);

export const Check: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <polyline points="20 6 9 17 4 12" />
    </BaseIcon>
);

export const Plus: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <line x1="12" y1="5" x2="12" y2="19" />
        <line x1="5" y1="12" x2="19" y2="12" />
    </BaseIcon>
);

export const Search: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <circle cx="11" cy="11" r="7" />
        <line x1="20" y1="20" x2="16.65" y2="16.65" />
    </BaseIcon>
);

export const Link: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
        <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" />
    </BaseIcon>
);

export const RefreshCw: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8" />
        <path d="M21 3v5h-5" />
        <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16" />
        <path d="M3 21v-5h5" />
    </BaseIcon>
);

export const Save: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
        <polyline points="17 21 17 13 7 13 7 21" />
        <polyline points="7 3 7 8 15 8" />
    </BaseIcon>
);

export const RotateCcw: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
        <path d="M3 3v5h5" />
    </BaseIcon>
);


export const Send: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <line x1="22" y1="2" x2="11" y2="13" />
        <polygon points="22 2 15 22 11 13 2 9 22 2" />
    </BaseIcon>
);

export const Microphone: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3Z" />
        <path d="M19 10v2a7 7 0 0 1-14 0v-2" />
        <line x1="12" y1="19" x2="12" y2="22" />
    </BaseIcon>
);
export const Trash: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M3 6h18" />
        <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6" />
        <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2" />
        <line x1="10" y1="11" x2="10" y2="17" />
        <line x1="14" y1="11" x2="14" y2="17" />
    </BaseIcon>
);
export const Download: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
        <polyline points="7 10 12 15 17 10" />
        <line x1="12" y1="15" x2="12" y2="3" />
    </BaseIcon>
);

export const Calendar: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <rect width="18" height="18" x="3" y="4" rx="2" ry="2" />
        <line x1="16" y1="2" x2="16" y2="6" />
        <line x1="8" y1="2" x2="8" y2="6" />
        <line x1="3" y1="10" x2="21" y2="10" />
    </BaseIcon>
);

export const Lock: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <rect width="18" height="11" x="3" y="11" rx="2" ry="2" />
        <path d="M7 11V7a5 5 0 0 1 10 0v4" />
    </BaseIcon>
);

export const Paperclip: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.51a2 2 0 0 1-2.83-2.83l8.49-8.48" />
    </BaseIcon>
);
export const AlertCircle: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <circle cx="12" cy="12" r="10" />
        <line x1="12" y1="8" x2="12" y2="12" />
        <line x1="12" y1="16" x2="12.01" y2="16" />
    </BaseIcon>
);

export const AlertTriangle: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z" />
        <line x1="12" y1="9" x2="12" y2="13" />
        <line x1="12" y1="17" x2="12.01" y2="17" />
    </BaseIcon>
);

export const MessageSquare: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
    </BaseIcon>
);

export const Upload: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
        <polyline points="17 8 12 3 7 8" />
        <line x1="12" y1="3" x2="12" y2="15" />
    </BaseIcon>
);

export const ArrowRight: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <line x1="5" y1="12" x2="19" y2="12" />
        <polyline points="12 5 19 12 12 19" />
    </BaseIcon>
);

export const ArrowUp: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <polyline points="18 15 12 9 6 15" />
    </BaseIcon>
);

export const ArrowDown: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <polyline points="6 9 12 15 18 9" />
    </BaseIcon>
);
export const User: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2" />
        <circle cx="12" cy="7" r="4" />
    </BaseIcon>
);

export const Target: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <circle cx="12" cy="12" r="10" />
        <circle cx="12" cy="12" r="6" />
        <circle cx="12" cy="12" r="2" />
    </BaseIcon>
);

export const Activity: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <polyline points="22 12 18 12 15 21 9 3 6 12 2 12" />
    </BaseIcon>
);

export const Clock: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <circle cx="12" cy="12" r="10" />
        <polyline points="12 6 12 12 16 14" />
    </BaseIcon>
);

export const AtSign: React.FC<IconProps> = (props) => (
    <BaseIcon {...props}>
        <circle cx="12" cy="12" r="4" />
        <path d="M16 8v5a3 3 0 0 0 6 0v-1a10 10 0 1 0-4 8" />
    </BaseIcon>
);
