import React from "react";
import { WorkItemPicker } from "./Common/WorkItemPicker";

// Dans votre composant:
const handleWorkItemSelected = (workItemId: number) => {
    // Mettre à jour votre state ou faire ce que vous voulez avec l'ID
    console.log(`Selected work item: ${workItemId}`);
};

// Dans votre render:
<WorkItemPicker onWorkItemSelected={handleWorkItemSelected} />
