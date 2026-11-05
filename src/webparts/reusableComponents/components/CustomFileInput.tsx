import * as React from 'react';
import { useRef } from 'react';
import { IIconProps, Label, PrimaryButton, } from '@fluentui/react';
import styles from './ReusableComponents.module.scss';
interface CustomFileInputProps {
  onChange: (event: React.ChangeEvent<HTMLInputElement>) => void;
  disabled?: boolean;
}
const CustomFileInput: React.FC<CustomFileInputProps> = ({ onChange, disabled }) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const chooseFileIcon: IIconProps = { iconName: "Attach" };
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    onChange(event);
  };
  // Trigger the hidden file input when the button is clicked
  const handleClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };
  return (
    <>
      <div>
        <Label>Upload Document :  </Label>
        <PrimaryButton onClick={handleClick} disabled={disabled} iconProps={chooseFileIcon} text={"Upload"} className={styles.uploadDiv} />
        {/* Multiple files can be uploaded at once. */}
        <input
          ref={fileInputRef}
          id="files"
          type="file"
          style={{ display: 'none', ...(disabled ? { cursor: "not-allowed", opacity: 0.5 } : {}) }}
          onChange={handleFileChange}
          multiple
          disabled={disabled}
          accept="
          audio/*,
          video/*,
          image/*,
          .doc, .docx, .pdf, .txt,
          .xls, .xlsx, application/vnd.ms-excel,
          .dwg, .dxf, .step, .stp, .igs, .iges, model/*"
        />
        {/* Single file can be uploaded at once.  */}
        {/* <input
          ref={fileInputRef}
          id="files"
          type="file"
          style={{ display: 'none', ...(disabled ? { cursor: "not-allowed", opacity: 0.5 } : {}) }}
          onChange={handleFileChange}
          multiple = {false}
          disabled={disabled}
          accept="
          audio/*,
          video/*,
          image/*,
          .doc, .docx, .pdf, .txt,
          .xls, .xlsx, application/vnd.ms-excel,
          .dwg, .dxf, .step, .stp, .igs, .iges, model/*"
        /> */}
      </div>
      <div style={{paddingTop:"10px"}}>    
        {disabled && (
          <span>
            Please select a category first
          </span>
        )}
        </div>
    </>
  );
};
export default CustomFileInput;
