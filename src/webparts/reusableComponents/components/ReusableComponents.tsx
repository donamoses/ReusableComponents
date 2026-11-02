import * as React from 'react';
import styles from './ReusableComponents.module.scss';
import type { IReusableComponentsProps } from '../interfaces/IReusableComponentsProps';
import { TextField } from '@fluentui/react';


export default class ReusableComponents extends React.Component<IReusableComponentsProps> {
  public render(): React.ReactElement<IReusableComponentsProps> {
    

    return (
      <section className={`${styles.container}`}>
        <div className={styles.formpopup}>
          <div className={styles.formheader}>
            <div className={styles.formtitle}>{this.props.wpproperties.WebpartTitle}</div>
          </div>
          <div className={styles.formbody}>
            {/* One Column */}
            <div className={styles.row}>
              <div className={styles.col12}>
                <TextField label="Column 1"  />
              </div>
            </div>
           {/* 3/4 column + 1/4 column */}
            <div className={styles.row}>
              <div className={styles.col8}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col4}>
                <TextField label="Column 2" value=""/>
              </div>
            </div>
            {/* 2 columns */}
            <div className={styles.row}>
              <div className={styles.col6}>
                <TextField label="Column 1" value=""/>
              </div>
               <div className={styles.col6}>
                <TextField label="Column 2" value=""/>
              </div>
            </div>
            {/* 3 columns */}
<div className={styles.row}>
              <div className={styles.col4}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col4}>
                <TextField label="Column 2" value=""  />
              </div>
              <div className={styles.col4}>
                <TextField label="Column 3" value=""/>
              </div>
            </div>
            {/* 4 columns */}
            <div className={styles.row}>
              <div className={styles.col2}>
                <TextField label="Column 1" value=""/>
              </div>
              <div className={styles.col2}>
                <TextField label="Column 2" value=""  />
              </div>
              <div className={styles.col2}>
                <TextField label="Column 3" value=""/>
              </div>
              <div className={styles.col2}>
                <TextField label="Column 4" value=""/>
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}

  
