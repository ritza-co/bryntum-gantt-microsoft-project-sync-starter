import { TaskModel } from '@bryntum/gantt';

// Custom event model
export default class CustomTaskModel extends TaskModel {
    static $name = 'CustomTaskModel';
    static fields = [
        { name : 'msdyn_outlinelevel', type : 'number' },
        { name : 'msdyn_displaysequence', type : 'number' }
    ];
    // disable percentDone editing
    isEditable(field) {
        return field !== 'percentDone' && super.isEditable(field);
    }
}
