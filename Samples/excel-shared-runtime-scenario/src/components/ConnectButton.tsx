import * as React from 'react';
import { Button } from 'office-ui-fabric-react';
import {insertValueInCell} from '../commands/commands';

const onClickFromDialogue = () => {
        const res = 'From Dialogue';
        Office.context.ui.messageParent(JSON.stringify({res}));
}

const onClickFromTaskPane = () => {
    insertValueInCell('From TaskPane')
}

const ConnectButton = () =>
            <div className='ms-welcome'>
                <div className='ms-welcome__main'>
                    <Button onClick={onClickFromDialogue}>Insert from dialogue</Button>
                    <br />
                    <Button onClick={onClickFromTaskPane}>Insert from taskpane</Button>
                </div>
            </div>;

export default ConnectButton;
