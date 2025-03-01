import './TestConfigurationPage.scss';
import * as React from 'react';
import { useState, useEffect, ChangeEvent } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { TextField, TextFieldWidth } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { Question, showRootComponent } from '../Common/Common';

const ConfigurationPage: React.FC = () => {
    const [questions, setQuestions] = useState<Question[]>([]);
    const [newQuestion, setNewQuestion] = useState('');
    const [newExpectedEntriesCount, setNewExpectedEntriesCount] = useState<number>(1);
    const [newLabels, setNewLabels] = useState<string[]>([]);
    const [newTypes, setNewTypes] = useState<string[]>([]);
    const [error, setError] = useState<string | null>(null);

    interface Question {
        id: string;
        text: string;
        expectedEntries: {
          count: number;
          labels: string[];
          types: string[];
          weights: number[]; // Ensure this is defined at the entry level
        };
      }
      
    useEffect(() => {
        initializePage();
    }, []);

    const initializePage = async () => {
        await SDK.init();
        await loadCurrentQuestions();
    };

    const loadCurrentQuestions = async () => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        const questionVersions = await dataManager.getValue<Question[][]>('questionVersions', { scopeType: 'Default' }) || [];
        const latestVersion = questionVersions[questionVersions.length - 1] || [];
        setQuestions(latestVersion);
    };

    const saveNewVersionOfQuestions = async (updatedQuestions: Question[]) => {
        try {
            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

            const questionVersions = await dataManager.getValue<Question[][]>('questionVersions', { scopeType: 'Default' }) || [];

            // Push the updated questions instead of the current state
            questionVersions.push([...updatedQuestions]);
            await dataManager.setValue('questionVersions', questionVersions, { scopeType: 'Default' });

            // Reload the latest questions after saving
            await loadCurrentQuestions();
        } catch (err) {
            console.error('Failed to save questions:', err);
            setError('Failed to save questions.');
        }
    };

    const validateQuestions = (questions: Question[]): boolean => {
        return Array.isArray(questions) && questions.every(q => q.id && q.text && q.expectedEntries);
    };

    const addQuestion = () => {
        if (newQuestion) {
            const entryCount = newExpectedEntriesCount;
            const newQuestionObj: Question = {
                id: Date.now().toString(),
                text: newQuestion,
                expectedEntries: {
                    count: entryCount,
                    labels: newLabels.slice(0, entryCount).map((label, i) => label || `Entry ${i + 1}`),
                    types: newTypes.slice(0, entryCount).map((type, i) => type || 'url'),
                    weights: Array.from({ length: entryCount }, (_, i) => Math.pow(2, i))
                }
            };

            setQuestions([...questions, newQuestionObj]);
            setNewQuestion('');
            setNewExpectedEntriesCount(1);
            setNewLabels([]);
            setNewTypes([]);
        }
    };

    const handleTextChange = (questionId: string, newText: string) => {
        setQuestions(questions.map(q =>
            q.id === questionId ? { ...q, text: newText } : q
        ));
    };

    const handleLabelChange = (questionId: string, index: number, newLabel: string) => {
        setQuestions(questions.map(q =>
            q.id === questionId ? {
                ...q,
                expectedEntries: {
                    ...q.expectedEntries,
                    labels: q.expectedEntries.labels.map((label, i) => i === index ? newLabel : label)
                }
            } : q
        ));
    };

    const handleTypeChange = (questionId: string, index: number, newType: string) => {
        setQuestions(questions.map(q =>
            q.id === questionId ? {
                ...q,
                expectedEntries: {
                    ...q.expectedEntries,
                    types: q.expectedEntries.types.map((type, i) => i === index ? newType : type)
                }
            } : q
        ));
    };

    const downloadQuestionnaireJson = () => {
        const json = JSON.stringify(questions, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'questionnaire.json'; // Download filename
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url); // Clean up the URL
    };

    const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>) => {
        event.persist(); // Ensure the event can be accessed asynchronously

        const file = event.target.files?.[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const result = e.target?.result;
                    const loadedQuestions = JSON.parse(result as string) as Question[];

                    // Display a validation alert if necessary
                    if (!validateQuestions(loadedQuestions)) {
                        alert('Invalid questionnaire format');
                        return;
                    }

                    // Correctly update the questions state with the uploaded data
                    setQuestions(loadedQuestions);

                    // Save as a new version using the updated questions state
                    await saveNewVersionOfQuestions(loadedQuestions);

                    alert('Questionnaire uploaded and saved as a new version successfully!');

                } catch (error) {
                    console.error('Error reading or parsing JSON file:', error);
                    alert('Failed to load and parse the JSON file');
                } finally {
                    event.target.value = ''; // Clear the file input
                }
            };
            reader.readAsText(file);
        }
    };

    return (
        <div className="configuration-container">
            <h2>Questionnaire Configuration</h2>

            <input type="file" accept="application/json" onChange={handleFileUpload} />
            <Button text="Download JSON" onClick={downloadQuestionnaireJson} />

            {questions.map(question => (
                <div className="question-item" key={question.id}>
                    <TextField
                        value={question.text || ''}
                        onChange={(e, value) => handleTextChange(question.id, value || '')}
                        width={TextFieldWidth.auto}
                        className="full-width-text-field"
                        placeholder="Edit question text"
                    />
                    <div>Entries:</div>
                    {question.expectedEntries.labels.map((label, index) => (
                        <div key={index}>
                            <TextField
                                value={label || ''}
                                onChange={(e, value) => handleLabelChange(question.id, index, value || '')}
                                className="entry-label"
                                placeholder={`Label for entry ${index + 1}`}
                            />
                            <select
                                value={question.expectedEntries.types[index] || 'url'}
                                onChange={(e) => handleTypeChange(question.id, index, e.target.value)}
                            >
                                <option value="url">URL</option>
                                <option value="boolean">Todo/Done</option>
                                <option value="workItem">Work Item</option>
                            </select>
                        </div>
                    ))}
                    <Button text="Remove" onClick={() => setQuestions(questions.filter(q => q.id !== question.id))} />
                </div>
            ))}

            <div className="new-question">
                <TextField
                    value={newQuestion}
                    onChange={(e, value) => setNewQuestion(value || '')}
                    placeholder="Enter new question"
                />
                <TextField
                    value={String(newExpectedEntriesCount)}
                    onChange={(e, value) => {
                        const numericValue = Number(value);
                        if (!isNaN(numericValue)) {
                            setNewExpectedEntriesCount(numericValue);
                        }
                    }}
                    placeholder="Number of entries"
                />
                {Array.from({ length: newExpectedEntriesCount }).map((_, index) => (
                    <div key={index}>
                        <TextField
                            value={newLabels[index] || ''}
                            onChange={(e, value) => {
                                setNewLabels(labels => {
                                    const updated = [...labels];
                                    updated[index] = value || '';
                                    return updated;
                                });
                            }}
                            placeholder={`Label for entry ${index + 1}`}
                        />
                        <select
                            value={newTypes[index] || 'url'}
                            onChange={(e) => {
                                setNewTypes(types => {
                                    const updated = [...types];
                                    updated[index] = e.target.value;
                                    return updated;
                                });
                            }}
                        >
                            <option value="url">URL</option>
                            <option value="boolean">Todo/Done</option>
                            <option value="workItem">Work Item</option>
                        </select>
                    </div>
                ))}
            </div>
            <Button text="Add Question" onClick={addQuestion} />
            <Button
                text="Save Questions"
                onClick={() => saveNewVersionOfQuestions(questions)}
            />

        </div>
    );
};

showRootComponent(<ConfigurationPage />, 'configuration-root');

export default ConfigurationPage;