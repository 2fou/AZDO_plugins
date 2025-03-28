import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { useState, useEffect } from 'react';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { showRootComponent } from '../Common/Common';

interface StoryConfig {
    title: string;
    tasks: string[];
}

interface WorkItemTypesConfig {
    epic: string;
    feature: string;
    story: string;
    task: string;
}

interface Config {
    featureTitle: string;
    stories: StoryConfig[];
    workItemTypes: WorkItemTypesConfig;
}

const SetupComponent: React.FC = () => {
    const [config, setConfig] = useState<Config>({
        featureTitle: '',
        stories: [],
        workItemTypes: {
            epic: 'Epic',
            feature: 'Feature',
            story: 'Product Backlog Item',
            task: 'Task'
        }
    });
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        SDK.init().then(async () => {
            try {
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
                const savedConfig = await dataManager.getValue<Config>('workItemConfig');

                if (savedConfig) {
                    // Ensure no undefined references
                    if (!savedConfig.workItemTypes) {
                        savedConfig.workItemTypes = {
                            epic: 'Epic',
                            feature: 'Feature',
                            story: 'Product Backlog Item',
                            task: 'Task'
                        };
                    }
                    setConfig(savedConfig);
                } else {
                    setConfig({
                        featureTitle: '',
                        stories: [],
                        workItemTypes: {
                            epic: 'Epic',
                            feature: 'Feature',
                            story: 'Product Backlog Item',
                            task: 'Task'
                        }
                    });
                }
            } catch (err) {
                setError('Failed to load configuration.');
            }
        });
    }, []);

    const handleSave = async () => {
        try {
            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

            const configToSave: Config = {
                ...config,
                workItemTypes: config.workItemTypes || {
                    epic: 'Epic',
                    feature: 'Feature',
                    story: 'Product Backlog Item',
                    task: 'Task'
                }
            };

            await dataManager.setValue('workItemConfig', configToSave);
            alert('Configuration saved!');
        } catch (err) {
            setError('Failed to save configuration.');
        }
    };

    const clearConfig = () => {
        setConfig({
            featureTitle: '',
            stories: [],
            workItemTypes: {
                epic: 'Epic',
                feature: 'Feature',
                story: 'Product Backlog Item',
                task: 'Task'
            }
        });
        alert('Configuration cleared. You can now start fresh.');
    };

    const addStory = () => {
        setConfig({
            ...config,
            stories: [...config.stories, { title: '', tasks: [''] }]
        });
    };

    const updateStoryTitle = (index: number, value: string) => {
        const newStories = [...config.stories];
        newStories[index].title = value;
        setConfig({ ...config, stories: newStories });
    };

    const updateTask = (storyIndex: number, taskIndex: number, value: string) => {
        const newStories = [...config.stories];
        newStories[storyIndex].tasks[taskIndex] = value;
        setConfig({ ...config, stories: newStories });
    };

    const addTask = (storyIndex: number) => {
        const newStories = [...config.stories];
        newStories[storyIndex].tasks.push('');
        setConfig({ ...config, stories: newStories });
    };

    // Helpers to update work item type names
    const updateWorkItemType = (field: keyof WorkItemTypesConfig, value: string) => {
        setConfig({
            ...config,
            workItemTypes: {
                ...config.workItemTypes,
                [field]: value
            }
        });
    };

    return (
        <div style={{ padding: '10px' }}>
            <h2>Work Item Configuration</h2>
            {error && <div className="error">{error}</div>}

            {/* Edit the names of each work item type */}
            <div style={{ marginLeft: '20px', marginTop: '10px', paddingBottom: '10px', borderBottom: '1px solid #ccc' }}>
                <h3>Work Item Types</h3>
                <div>
                    <TextField
                        label="Epic Type Name"
                        value={config.workItemTypes.epic}
                        onChange={(e, val) => updateWorkItemType('epic', val || '')}
                    />
                    <TextField
                        label="Feature Type Name"
                        value={config.workItemTypes.feature}
                        onChange={(e, val) => updateWorkItemType('feature', val || '')}
                    />
                    <TextField
                        label="Story Type Name"
                        value={config.workItemTypes.story}
                        onChange={(e, val) => updateWorkItemType('story', val || '')}
                    />
                    <TextField
                        label="Task Type Name"
                        value={config.workItemTypes.task}
                        onChange={(e, val) => updateWorkItemType('task', val || '')}
                    />
                </div>
            </div>

            {/* Feature section */}
            <div style={{ marginLeft: '20px', marginTop: '20px', paddingBottom: '10px', borderBottom: '1px solid #ccc' }}>
                <strong>{config.workItemTypes.epic}</strong>
                <div style={{ marginLeft: '20px' }}>
                    <TextField
                        value={config.featureTitle}
                        onChange={(e, val) => setConfig({ ...config, featureTitle: val || '' })}
                        placeholder={`Enter a ${config.workItemTypes.feature} title`}
                    />

                    <Button text={`Add ${config.workItemTypes.story}`} onClick={addStory} />

                    {config.stories.map((story, storyIndex) => (
                        <div key={`story-${storyIndex}`} style={{ marginLeft: '20px', marginTop: '10px' }}>
                            <TextField
                                value={story.title}
                                onChange={(e, val) => updateStoryTitle(storyIndex, val || '')}
                                placeholder={`${config.workItemTypes.story} ${storyIndex + 1} Title`}
                            />

                            {story.tasks.map((task, taskIndex) => (
                                <div key={`task-${taskIndex}`} style={{ marginLeft: '20px' }}>
                                    <TextField
                                        value={task}
                                        onChange={(e, val) => updateTask(storyIndex, taskIndex, val || '')}
                                        placeholder={`${config.workItemTypes.task} ${taskIndex + 1}`}
                                    />
                                </div>
                            ))}

                            <Button
                                text={`Add ${config.workItemTypes.task}`}
                                onClick={() => addTask(storyIndex)}
                                style={{ marginLeft: '20px' }}
                            />
                        </div>
                    ))}
                </div>
            </div>

            <Button
                text="Save Configuration"
                onClick={handleSave}
                primary={true}
                style={{ marginTop: '20px' }}
            />
            <Button
                text="Clear Configuration"
                onClick={clearConfig}
                style={{ marginTop: '10px' }}
            />
        </div>
    );
};

showRootComponent(<SetupComponent />, 'setupcomponent-root');
export default SetupComponent;