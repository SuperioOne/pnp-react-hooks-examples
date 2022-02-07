import { IDropdownOption, Stack } from "@fluentui/react";
import { useSubWebInfos } from "pnp-react-hooks";
import * as React from "react";

export function ListInspector()
{
    const subSitesOptions: IDropdownOption[] = useSubWebInfos({
        query: {
            select: ["Title", "ServerRelativeUrl"]
        },
        useCache: true
    })
        .map(e => ({ key: e.ServerRelativeUrl, text: e.Title }));





    return (
        <div>
            <Stack>
                <Stack.Item>

                </Stack.Item>
            </Stack>
        </div>);
}