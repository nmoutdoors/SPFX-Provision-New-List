import * as React from 'react';
import styles from './ListProvisionTest.module.scss';
import type { IListProvisionTestProps } from './IListProvisionTestProps';
import { Text, PrimaryButton } from '@fluentui/react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn
} from '@fluentui/react/lib/DetailsList';

export default class ListProvisionTest extends React.Component<IListProvisionTestProps, {}> {
  private _columns: IColumn[] = [
    {
      key: 'id',
      name: 'ID',
      fieldName: 'ID',
      minWidth: 50,
      maxWidth: 50,
    },
    {
      key: 'title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 200,
      maxWidth: 300,
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 100,
      maxWidth: 100,
    },
    {
      key: 'assignedTo',
      name: 'Assigned To',
      fieldName: 'AssignedTo',
      minWidth: 150,
      maxWidth: 200,
      onRender: (item) => item.AssignedTo ? item.AssignedTo.Title : ''
    }
  ];

  public render(): React.ReactElement<IListProvisionTestProps> {
    const {
      projectsListExists,
      onConfigureClick,
      userDisplayName,
      items
    } = this.props;

    return (
      <section className={styles.listProvisionTest}>
        {!projectsListExists ? (
          <div>
            <Text variant="xLarge" block>
              Welcome, {userDisplayName}!
            </Text>
            <h4>Let&apos;s get started configuring your new webpart. Click the Open Settings button</h4>
            <PrimaryButton 
              text="Open Settings" 
              onClick={onConfigureClick}
              styles={{ root: { marginTop: '10px' } }}
            />
          </div>
        ) : (
          <div>
            <Text variant="xLarge" block style={{ marginBottom: '20px' }}>
              Projects List Items
            </Text>
            <DetailsList
              items={items || []}
              columns={this._columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
            />
          </div>
        )}
      </section>
    );
  }
}

