import * as React from 'react';
import styles from './DocumentManagementInterface.module.scss';
import type { IDocumentManagementInterfaceProps } from './IDocumentManagementInterfaceProps';
import { Sidebar } from './Sidebar/Sidebar';
import { Projects } from './Projects/Projects';
import { Documents } from './Documents/Documents';
import { Transmittals } from './Transmittals/Transmittals';
import { DistributionLists } from './DistributionLists/DistributionLists';
import { DocumentHistory } from './DocumentHistory/DocumentHistory';
import { Alerts } from './Alerts/Alerts';
import { Files } from './Files/Files';
import { Icon } from '@fluentui/react/lib/Icon';

interface IDocumentManagementInterfaceState {
  selectedTab: string;
  isSidebarOpen: boolean;
  isMobile: boolean;
}

export default class DocumentManagementInterface extends React.Component<
  IDocumentManagementInterfaceProps,
  IDocumentManagementInterfaceState
> {
  private resizeTimeout: number | undefined;

  constructor(props: IDocumentManagementInterfaceProps) {
    super(props);
    this.state = {
      selectedTab: 'projects',
      isSidebarOpen: false,
      isMobile: this.checkIsMobile(),
    };
  }

  public componentDidMount(): void {
    window.addEventListener('resize', this.handleResize);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('resize', this.handleResize);
    if (this.resizeTimeout) {
      window.clearTimeout(this.resizeTimeout);
    }
  }

  private checkIsMobile = (): boolean => {
    return window.innerWidth < 768;
  };

  private handleResize = (): void => {
    if (this.resizeTimeout) {
      window.clearTimeout(this.resizeTimeout);
    }

    this.resizeTimeout = window.setTimeout(() => {
      const isMobile = this.checkIsMobile();
      this.setState({ isMobile });

      // Chiudi la sidebar su mobile quando si ridimensiona a desktop
      if (!isMobile && this.state.isSidebarOpen) {
        this.setState({ isSidebarOpen: false });
      }
    }, 150);
  };

  private handleTabChange = (tabId: string): void => {
    this.setState({ selectedTab: tabId });

    // Su mobile, chiudi la sidebar dopo aver selezionato una tab
    if (this.state.isMobile) {
      this.setState({ isSidebarOpen: false });
    }
  };

  private toggleSidebar = (): void => {
    this.setState({ isSidebarOpen: !this.state.isSidebarOpen });
  };

  private closeSidebar = (): void => {
    this.setState({ isSidebarOpen: false });
  };

  private renderTabContent(): React.ReactElement {
    switch (this.state.selectedTab) {
      case 'projects':
        return <Projects context={this.props.context} />;
      case 'documents':
        return <Documents context={this.props.context} />;
      case 'transmittals':
        return <Transmittals context={this.props.context} />;
      case 'distributionLists':
        return <DistributionLists context={this.props.context} />;
      case 'documentHistory':
        return <DocumentHistory context={this.props.context} />;
      case 'alerts':
        return <Alerts context={this.props.context} />;
      case 'files':
        return <Files context={this.props.context} />;
      default:
        return <div>Seleziona una tab</div>;
    }
  }

  public render(): React.ReactElement<IDocumentManagementInterfaceProps> {
    const { selectedTab, isSidebarOpen, isMobile } = this.state;

    return (
      <section className={styles.documentManagementInterface}>
        {/* Mobile Overlay */}
        {isMobile && isSidebarOpen && (
          <div
            className={`${styles.mobileOverlay} ${isSidebarOpen ? styles.visible : ''}`}
            onClick={this.closeSidebar}
          />
        )}

        {/* Sidebar */}
        <Sidebar
          selectedTab={selectedTab}
          onTabChange={this.handleTabChange}
          isVisible={!isMobile || isSidebarOpen}
          onClose={this.closeSidebar}
        />

        {/* Main Content */}
        <div className={styles.mainContent}>{this.renderTabContent()}</div>

        {/* Mobile Menu Toggle */}
        {isMobile && (
          <button
            className={styles.mobileMenuToggle}
            onClick={this.toggleSidebar}
            aria-label="Toggle menu"
          >
            <Icon iconName={isSidebarOpen ? 'Cancel' : 'GlobalNavButton'} />
          </button>
        )}
      </section>
    );
  }
}