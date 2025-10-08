import * as React from 'react';
import type { ISmalsusPurchasesProps } from './ISmalsusPurchasesProps';
import App from './App';
import { TailwindLoader } from '../utils/tailwindLoader';
import './Style.css';

export default class SmalsusPurchases extends React.Component<ISmalsusPurchasesProps> {
  public componentDidMount(): void {
    // Load Tailwind CSS CDN and override CanvasComponent styles
    TailwindLoader.loadTailwind();

    // Also manually trigger CanvasComponent style override in case it's needed immediately
    TailwindLoader.overrideCanvasComponent();
  }

  public render(): React.ReactElement<ISmalsusPurchasesProps> {
    return (
      <div style={{ width: '100%', height: '100%' }}>
        <App context={this.props.context} />
      </div>
    );
  }
}
