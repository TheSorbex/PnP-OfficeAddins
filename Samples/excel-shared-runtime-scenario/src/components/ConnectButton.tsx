import * as React from 'react';
import { Button } from 'office-ui-fabric-react';
const onClickFromDialogue = () => {
        const res = 'From Dialogue';
        Office.context.ui.messageParent(JSON.stringify({res}));
}
const ConnectButton = () =>
            <div className='ms-welcome'>
                <div className='ms-welcome__main'>
                    <Button onClick={onClickFromDialogue}>Insert from dialogue</Button>
                </div>
            </div>;
export default ConnectButton;
