import * as React from 'react';
import { useState, useEffect } from 'react';
import { Button } from 'azure-devops-ui/Button';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import * as SDK from 'azure-devops-extension-sdk';
import { showRootComponent, Deliverable } from '../Common/Common';



const DeliverableConfigurationPage: React.FC = () => {
    const [deliverables, setDeliverables] = useState<Deliverable[]>([]);
    const [newDeliverable, setNewDeliverable] = useState<Deliverable>({
        id: '',
        label: '',
        type: 'url',
        value: '',
    });

    useEffect(() => {
        initializePage();
    }, []);

    const initializePage = async () => {
        await SDK.init();
        await loadDeliverables();
    };

    const saveDeliverables = async (updatedDeliverables: Deliverable[]) => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        await dataManager.setValue('deliverables', updatedDeliverables, { scopeType: 'Default' });
    };

    const addDeliverable = () => {
        if (newDeliverable.label) {
            const updatedDeliverables = [...deliverables, { ...newDeliverable, id: Date.now().toString(), linkedQuestions: [] }];
            setDeliverables(updatedDeliverables);
            setNewDeliverable({ id: '', label: '', type: 'url', value: '' });
            saveDeliverables(updatedDeliverables);
        } else {
            alert("Deliverable label is required!");
        }
    };

    const loadDeliverables = async () => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        const storedDeliverables = (await dataManager.getValue<Deliverable[]>('deliverables', { scopeType: 'Default' }) || []).map(d => ({
            ...d,
        }));
        setDeliverables(storedDeliverables);
    };

    const updateDeliverable = (deliverableId: string, updated: Partial<Deliverable>) => {
        const updatedDeliverables = deliverables.map(del =>
            del.id === deliverableId ? { ...del, ...updated } : del
        );
        setDeliverables(updatedDeliverables);
        saveDeliverables(updatedDeliverables);
    };

    const deleteDeliverable = (deliverableId: string) => {
        const updatedDeliverables = deliverables.filter(deliverable => deliverable.id !== deliverableId);
        setDeliverables(updatedDeliverables);
        saveDeliverables(updatedDeliverables);
    };



    // Code in the main component render
    return (
        <div className="deliverable-configuration-container" style={{ maxWidth: '800px', margin: '0 auto', padding: '20px' }}>
            <h2>Deliverable Configuration</h2>
            <div className="deliverable-form">
                <TextField
                    value={newDeliverable.label}
                    onChange={(e, value) => setNewDeliverable({ ...newDeliverable, label: value || '' })}
                    placeholder="Deliverable Label (required)"
                    width={TextFieldWidth.standard}
                />
                <select
                    value={newDeliverable.type}
                    onChange={(e) => setNewDeliverable({ ...newDeliverable, type: e.target.value })}
                    style={{ marginBottom: '10px' }}
                >
                    <option value="url">URL</option>
                    <option value="boolean">Todo/Done</option>
                    <option value="workItem">Work Item</option>
                </select>
                <Button text="Add Deliverable" onClick={addDeliverable} />
            </div>
            <ul className="deliverable-list" style={{ listStyleType: 'none', padding: 0 }}>
                {deliverables.map(deliverable => (
                    <li key={deliverable.id} style={{ marginBottom: '10px' }}>
                        <TextField
                            value={deliverable.label}
                            onChange={(e, value) => updateDeliverable(deliverable.id, { label: value || '' })}
                            width={TextFieldWidth.standard}
                        />
                        <select
                            value={deliverable.type}
                            onChange={(e) => updateDeliverable(deliverable.id, { type: e.target.value })}
                        >
                            <option value="url">URL</option>
                            <option value="boolean">Todo/Done</option>
                            <option value="workItem">Work Item</option>
                        </select>
                        <Button text="Delete" onClick={() => deleteDeliverable(deliverable.id)} />
                    </li>
                ))}
            </ul>
        </div>
    );
};

showRootComponent(<DeliverableConfigurationPage />, 'deliverable-configuration-root');

export default DeliverableConfigurationPage;