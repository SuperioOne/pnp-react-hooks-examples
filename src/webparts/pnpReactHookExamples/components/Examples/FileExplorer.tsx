import * as React from "react";
import { DefaultButton, Stack } from "@fluentui/react";
import { useFolderTree, useWebInfo } from "pnp-react-hooks";

export function FileExplorer()
{
    const webInfo = useWebInfo();
    const treeContext = useFolderTree(webInfo?.ServerRelativeUrl);

    React.useEffect(() =>
    {
        console.debug(treeContext);
    }, [treeContext]);

    return (
        <Stack>
            <Stack.Item>
                <DefaultButton text="Home" onClick={() => treeContext?.home?.()} />
            </Stack.Item>
            <Stack.Item>
                <DefaultButton text="Up" onClick={() => treeContext?.up?.()} />
            </Stack.Item>
            <Stack.Item>
                <DefaultButton text="Next" onClick={() => treeContext?.folders[0].setAsRoot()} />
            </Stack.Item>
        </Stack>
    );
}