import * as React from 'react';
import {
    Modal,
    Spinner,
    SpinnerSize,
    mergeStyleSets
} from '@fluentui/react';

export interface IModalOverlayProps {
    isModalOpen: boolean;
    modalText: string;
}

export interface IModalOverlayState {
    isModalOpen: boolean;
    modalText: string;
}

export default class ModalOverlay extends React.Component<IModalOverlayProps, IModalOverlayState> {
    constructor(props: IModalOverlayProps) {
        super(props);
        this.state = {
            isModalOpen: props.isModalOpen,
            modalText: props.modalText
        }
    }

    componentDidMount(): void {
        this.setState({ isModalOpen: this.props.isModalOpen, modalText: this.props.modalText });
    }

    componentDidUpdate(prevProps: Readonly<IModalOverlayProps>): void {
        if (prevProps !== this.props) {
            this.setState({ isModalOpen: this.props.isModalOpen, modalText: this.props.modalText });
        }
    }

    private closeModal() {
        this.setState({ isModalOpen: false, modalText: "" })
    }

    public render(): React.ReactElement<IModalOverlayProps> {

        const contentStyles = mergeStyleSets({
            container: {
                display: 'flex',
                flexFlow: 'column nowrap',
                alignItems: 'stretch',
            },
            body: {
                padding: '5em 0em'
            },
        });

        return (
            <div>
                <Modal
                    isModeless={false}
                    isOpen={this.state.isModalOpen}
                    onDismiss={this.closeModal}
                    isBlocking={true}
                    containerClassName={contentStyles.container}
                >
                    <div className={contentStyles.body}>
                        <Spinner label={this.state.modalText} size={SpinnerSize.large} />
                    </div>
                </Modal>
            </div>
        );
    }
}
