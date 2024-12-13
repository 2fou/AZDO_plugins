{
  "questions": [
    {
      "id": "q1",
      "text": "Is the feature fully tested?",
      "type": "checkbox"
    },
    {
      "id": "q2",
      "text": "Are there any known issues?",
      "type": "checkbox"
    }
  ]
}
tfx extension publish --manifest-globs azure-devops-extension.json --overrides-file configs/dev.json --token [token]
