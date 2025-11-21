import * as React from 'react';
import styles from './DocumentManagementInterface.module.scss';
import type { IDocumentManagementInterfaceProps } from './IDocumentManagementInterfaceProps';
import { Sidebar } from './Sidebar/Sidebar';
import { Projects } from './Projects/Projects';
import { Documents } from './Documents/Documents';
import { Transmittals } from './Transmittals/Transmittals';
import { DistributionLists } from './DistributionLists/DistributionLists';
import { Files } from './Files/Files';

export default class DocumentManagementInterface extends React.Component<IDocumentManagementInterfaceProps, { selectedTab: string }> {
  constructor(props: IDocumentManagementInterfaceProps) {
    super(props);
    this.state = {
      selectedTab: 'projects', // default tab
    };
  }

  handleTabChange = (tabId: string) => {
    this.setState({ selectedTab: tabId });
  };

  renderTabContent() {
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
        return <div>Document History</div>;
      case 'alerts':
        return <div>Alerts</div>;
      case 'files':
        return <Files context={this.props.context} />;

      default:
        return <div>Seleziona una tab</div>;
    }
  }

  public render(): React.ReactElement<IDocumentManagementInterfaceProps> {
    return (
      <section className={styles.documentManagementInterface}>
        <Sidebar selectedTab={this.state.selectedTab} onTabChange={this.handleTabChange} isVisible={true} />
        <div className={styles.mainContent}>
          {this.renderTabContent()}
        </div>
      </section>
    );
  }
}
