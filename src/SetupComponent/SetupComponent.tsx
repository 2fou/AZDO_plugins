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

interface Config {
    featureTitle: string;
    stories: StoryConfig[];
}

const SetupComponent: React.FC = () => {
    const [config, setConfig] = useState<Config>({
        featureTitle: '',
        stories: []
    });
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        SDK.init().then(async () => {
            try {
                const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
                const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
                const savedConfig = await dataManager.getValue<Config>('workItemConfig');
                if (savedConfig && Array.isArray(savedConfig.stories)) {
                    setConfig(savedConfig);
                } else {
                    setConfig({ featureTitle: '', stories: [] });
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
            await dataManager.setValue('workItemConfig', config);
            alert('Configuration saved!');
        } catch (err) {
            setError('Failed to save configuration.');
        }
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

    return (
        <div style={{ padding: '10px' }}>
            <h2>Work Item Configuration</h2>
            {error && <div className="error">{error}</div>}
            
            <TextField
                value={config.featureTitle}
                onChange={(e, val) => setConfig({ ...config, featureTitle: val || '' })}
                placeholder="Feature Title"
            />
            
            <Button text="Add Story" onClick={addStory} />
            
            {Array.isArray(config.stories) && config.stories.map((story, storyIndex) => (
                <div key={`story-${storyIndex}`} style={{ marginLeft: '20px', marginTop: '10px' }}>
                    <TextField
                        value={story.title}
                        onChange={(e, val) => updateStoryTitle(storyIndex, val || '')}
                        placeholder={`Story ${storyIndex + 1} Title`}
                    />
                    
                    {story.tasks.map((task, taskIndex) => (
                        <div key={`task-${taskIndex}`} style={{ marginLeft: '20px' }}>
                            <TextField
                                value={task}
                                onChange={(e, val) => updateTask(storyIndex, taskIndex, val || '')}
                                placeholder={`Task ${taskIndex + 1}`}
                            />
                        </div>
                    ))}
                    
                    <Button 
                        text="Add Task" 
                        onClick={() => addTask(storyIndex)}
                        style={{ marginLeft: '20px' }}
                    />
                </div>
            ))}
            
            <Button 
                text="Save Configuration" 
                onClick={handleSave} 
                primary={true} 
                style={{ marginTop: '20px' }}
            />
        </div>
    );
};

showRootComponent(<SetupComponent />, 'setupcomponent-root');
export default SetupComponent;