import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { useState, useEffect } from 'react';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { showRootComponent } from '../Common/Common';

interface Config {
    featureTitle: string;
    storyTitle: string;
    taskTitles: string[];
}

const SetupComponent: React.FC = () => {
    const [config, setConfig] = useState<Config>({
        featureTitle: '',
        storyTitle: '',
        taskTitles: ['', '']
    });
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        console.log('Initializing SDK...');
        SDK.init().then(async () => {
            console.log('SDK initialized.');
            // Load initial data
            try {
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
                console.log('Fetching configuration...');
                const savedConfig: Config = await dataManager.getValue<Config>('workItemConfig', { defaultValue: config, scopeType: 'Default' });
                console.log('Configuration loaded:', savedConfig);
                setConfig(savedConfig);
            } catch (err) {
                console.error('Failed to load configuration:', err);
                setError('Failed to load configuration.');
            }
        }).catch(err => {
            console.error('Failed to initialize SDK:', err);
            setError('Failed to initialize SDK.');
        });
    }, []);

    const handleSave = async () => {
        try {
            console.log('Saving configuration...', config);
            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
            await dataManager.setValue('workItemConfig', { defaultValue: config, scopeType: 'Default' });
            console.log('Configuration saved:', config);
            alert('Configuration saved successfully!');
        } catch (err) {
            console.error('Failed to save configuration:', err);
            setError('Failed to save configuration.');
        }
    };

    return (
        <div>
            <h2>Setup Configuration</h2>
            {error && <div className="error">{error}</div>}
            <TextField
                value={config.featureTitle}
                onChange={(e, value) => setConfig({ ...config, featureTitle: value || '' })}
                placeholder="Feature Title"
            />
            <TextField
                value={config.storyTitle}
                onChange={(e, value) => setConfig({ ...config, storyTitle: value || '' })}
                placeholder="Story Title"
            />
            {config.taskTitles.map((task, index) => (
                <TextField
                    key={`task-${index}`}
                    value={task}
                    onChange={(e, value) => {
                        const newTaskTitles = [...config.taskTitles];
                        newTaskTitles[index] = value || '';
                        setConfig({ ...config, taskTitles: newTaskTitles });
                    }}
                    placeholder={`Task ${index + 1} Title`}
                />
            ))}
            <Button text="Save Configuration" onClick={handleSave} />
        </div>
    );
};
showRootComponent(<SetupComponent />, 'setupcomponent-root');
export default SetupComponent;