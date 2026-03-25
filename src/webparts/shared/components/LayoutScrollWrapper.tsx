import * as React from 'react';
import './LayoutScrollWrapper.css';

export interface ILayoutScrollWrapperProps {
    children: React.ReactNode;
    className?: string;
    innerClassName?: string;
    style?: React.CSSProperties;
}

export default function LayoutScrollWrapper(props: ILayoutScrollWrapperProps): JSX.Element {
    const { children, className = '', innerClassName = '', style } = props;

    return (
        <div className={`layout-scroll-wrapper ${className}`.trim()} style={style}>
            <div className={`layout-scroll-wrapper__inner ${innerClassName}`.trim()}>
                {children}
            </div>
        </div>
    );
}
