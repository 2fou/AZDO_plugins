import * as React from 'react';
import { useState, useEffect } from 'react';
import * as SDK from 'azure-devops-extension-sdk';
import { Button } from 'azure-devops-ui/Button';
import { TextField } from 'azure-devops-ui/TextField';
import { IExtensionDataService, CommonServiceIds } from 'azure-devops-extension-api';
import { Question, normalizeQuestions, showRootComponent } from '../Common/Common';

const ConfigurationPage: React.FC = () => {
    const [questions, setQuestions] = useState<Question[]>([]);
    const [newQuestion, setNewQuestion] = useState('');
    const [newExpectedEntriesCount, setNewExpectedEntriesCount] = useState<number>(1);
    const [newLabels, setNewLabels] = useState<string[]>([]);
    const [newTypes, setNewTypes] = useState<string[]>([]);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        const initializeSDK = async () => {
            console.log('Initializing SDK...');
            await SDK.init();
            try {
                await loadQuestions();
            } catch (err) {
                console.error('Failed to load questions:', err);
                setError('Failed to load questions.');
            } finally {
                setIsLoading(false);
                console.log('Initialization complete.');
            }
        };

        initializeSDK();
    }, []);

    const loadQuestions = async () => {
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

        const loadedQuestionsRaw = await dataManager.getValue<Question[]>('questions', { scopeType: 'Default' }) || [];
        console.log('Questions loaded (raw):', loadedQuestionsRaw);
        const loadedQuestions = normalizeQuestions(loadedQuestionsRaw);
        setQuestions(loadedQuestions);
        console.log('Questions loaded:', loadedQuestions);
    };

    const saveQuestions = async () => {
        try {
            const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
            const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

            await dataManager.setValue('questions', questions, { scopeType: 'Default' });
            alert('Questions saved successfully!');
        } catch (err) {
            console.error('Failed to save questions:', err);
            setError('Failed to save questions.');
        }
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
                    types: newTypes.slice(0, entryCount).map((type, i) => type || 'url'), // Ensure it always defaults to 'url'
                    weights: Array.from({ length: entryCount }, (_, i) => Math.pow(2, i))
                }
            };

            setQuestions([...questions, newQuestionObj]);
            setNewQuestion('');
            setNewExpectedEntriesCount(1);
            setNewLabels([]);
            setNewTypes([]);  // Reset types
        }
    };

    const handleExpectedEntriesCountChange = (questionId: string, count: number) => {
        setQuestions(questions.map(q =>
            q.id === questionId
                ? { ...q, expectedEntries: { ...q.expectedEntries, count } }
                : q
        ));
    };

    const handleLabelChange = (questionId: string, index: number, label: string) => {
        setQuestions(questions.map(q =>
            q.id === questionId
                ? {
                    ...q,
                    expectedEntries: {
                        ...q.expectedEntries,
                        labels: q.expectedEntries.labels.map((l, i) => i === index ? label : l)
                    }
                }
                : q
        ));
    };

    const handleTypeChange = (questionId: string, index: number, type: string) => {
        setQuestions(questions.map(q =>
            q.id === questionId
                ? {
                    ...q,
                    expectedEntries: {
                        ...q.expectedEntries,
                        types: q.expectedEntries.types.map((t, i) => i === index ? type : t)
                    }
                }
                : q
        ));
    };

    return (
        <div>
            <h2>Questionnaire Configuration</h2>
            {questions.map(question => (
                <div key={question.id}>
                    <TextField value={question.text} readOnly />
                    <TextField
                        value={question.expectedEntries.count.toString()}
                        onChange={(e, value) => {
                            const numericValue = Number(value);
                            if (!isNaN(numericValue)) {
                                handleExpectedEntriesCountChange(question.id, numericValue);
                            }
                        }}
                        placeholder="Number of entries"
                    />
                    {Array.from({ length: question.expectedEntries.count }).map((_, index) => (
                        <div key={index}>
                            <TextField
                                value={question.expectedEntries.labels[index] || ''}
                                onChange={(e, value) => handleLabelChange(question.id, index, value || '')}
                                placeholder={`Label for entry ${index + 1}`}
                            />
                            <select 
                                value={question.expectedEntries.types[index] || 'url'} // Reflect saved type
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
            <TextField
                value={newQuestion}
                onChange={(e, value) => setNewQuestion(value)}
                placeholder="Enter new question"
            />
            <TextField
                value={newExpectedEntriesCount.toString()}
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
                        onChange={(e, value) => setNewLabels(labels => {
                            const updated = [...labels];
                            updated[index] = value || '';
                            return updated;
                        })}
                        placeholder={`Label for entry ${index + 1}`}
                    />
                    <select 
                        value={newTypes[index] || 'url'} // Default for new entries
                        onChange={(e) => setNewTypes(types => {
                            const updated = [...types];
                            updated[index] = e.target.value;
                            return updated;
                        })}
                    >
                        <option value="url">URL</option>
                        <option value="boolean">Todo/Done</option>
                        <option value="workItem">Work Item</option>
                    </select>
                </div>
            ))}
            <Button text="Add Question" onClick={addQuestion} />
            <Button text="Save Questions" onClick={saveQuestions} />
        </div>
    );
};

showRootComponent(<ConfigurationPage />, 'configuration-root')
export default ConfigurationPage;