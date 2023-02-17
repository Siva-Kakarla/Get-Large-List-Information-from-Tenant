import * as React from 'react';
import {Modal} from 'react-bootstrap';

export interface IPopUpCompProps{
    message: string;
    childFunc():any;
}

export interface IPopUpCompState{
    _show: boolean;
}

export default class PopUpComp extends React.Component<IPopUpCompProps, IPopUpCompState> {

    constructor(props: IPopUpCompProps)
    {
        super(props);
        this.setState({
            _show:true
        });
    }

    public handleClose = () => {
        this.setState({
            _show:false
        });
    };

    public render() : React.ReactElement<IPopUpCompProps> {

        return (
            <div className="modal show" style={{ display: 'block', position: 'initial' }}>
                <Modal
                    show={this.state._show}
                    onHide={this.handleClose}
                    backdrop="static"
                    keyboard={false}
                    size="lg"
                    aria-labelledby="contained-modal-title-vcenter"
                    centered>

                <Modal.Header>
                    <Modal.Title>Modal title</Modal.Title>
                </Modal.Header>
        
                <Modal.Body>
                    <p>this.props.message</p>
                </Modal.Body>
        
                <Modal.Footer>
                    
                </Modal.Footer>
                </Modal>
            </div>
        );
    }
}