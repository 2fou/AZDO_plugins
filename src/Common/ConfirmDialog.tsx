import React from 'react';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

interface ConfirmDialogProps {
    title: string;
    message: string;
    onConfirm: () => void;
    onCancel: () => void;
    visible: boolean;
}

const ConfirmDialog: React.FC<ConfirmDialogProps> = ({ title, message, onConfirm, onCancel, visible }) => {
    return (
        <Dialog
            hidden={!visible}
            onDismiss={onCancel}
            dialogContentProps={{
                title: title,
                subText: message,
            }}
            modalProps={{
                isBlocking: true,
            }}
        >
            <DialogFooter>
                <PrimaryButton onClick={onConfirm} text="Confirm" />
                <DefaultButton onClick={onCancel} text="Cancel" />
            </DialogFooter>
        </Dialog>
    );
};

export default ConfirmDialog;