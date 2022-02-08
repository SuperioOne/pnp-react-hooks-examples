import * as React from "react";
import { DefaultButton, Dropdown, IDropdownOption, Separator, Stack, Text, Toggle } from "@fluentui/react";
import { useList, useLists, useSubWebInfos } from "pnp-react-hooks";
import AceEditor from "react-ace";

export function ListInfoInspector()
{
    const [selectedSite, setSelectedSite] = React.useState<string>(null);
    const [selectedList, setSelectedList] = React.useState<string>(null);
    const [showHidden, setShowHidden] = React.useState<boolean>(false);

    const siteAddress = selectedSite
        ? `${window.location.origin}${selectedSite}`
        : undefined;

    const subSiteOptions: IDropdownOption[] = useSubWebInfos({
        query: {
            select: ["Title", "ServerRelativeUrl"]
        },
        useCache: true
    })?.map(e => ({ key: e.ServerRelativeUrl, text: e.Title }));

    const listsOptions: IDropdownOption[] = useLists({
        query: {
            select: ["Title", "Id"],
            filter: showHidden ? undefined : "Hidden eq false",
            orderBy: "Title",
            orderyByAscending: true
        },
        web: siteAddress
    })?.map(e => ({ key: e.Id, text: e.Title }));

    const listInfo = useList(selectedList, {
        web: siteAddress
    });

    return (
        <div>
            <Stack tokens={{ childrenGap: 20 }}>
                <Stack.Item>
                    <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 15 }}>
                        <Stack.Item>
                            <Dropdown
                                disabled={!Array.isArray(subSiteOptions)}
                                label="Sub Sites"
                                options={subSiteOptions}
                                placeholder="(optional) Select a sub site"
                                selectedKey={selectedSite}
                                styles={dropdownStyles}
                                onChange={(_, opt) =>
                                {
                                    setSelectedSite(opt.key?.toString());
                                    setSelectedList(null);
                                }}
                            />
                        </Stack.Item>
                        <Stack.Item>
                            <DefaultButton
                                text="Clear"
                                onClick={() =>
                                {
                                    setSelectedSite(null);
                                    setSelectedList(null);
                                }}
                            />
                        </Stack.Item>
                    </Stack>
                </Stack.Item>
                <Stack.Item>
                    <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 15 }}>
                        <Stack.Item>
                            <Dropdown
                                options={listsOptions}
                                disabled={!Array.isArray(listsOptions)}
                                styles={dropdownStyles}
                                label="Lists"
                                placeholder="Select a list"
                                selectedKey={selectedList}
                                onChange={(_, opt) => setSelectedList(opt.key?.toString())}
                            />
                        </Stack.Item>
                        <Stack.Item>
                            <Toggle
                                label="Show hidden lists"
                                onText="Yes"
                                offText="No"
                                checked={showHidden}
                                onChange={(_, checked) =>
                                {
                                    setShowHidden(checked);
                                    setSelectedList(null);
                                }}
                            />
                        </Stack.Item>
                    </Stack>
                </Stack.Item>
                <Stack.Item >
                    {
                        listInfo && selectedList
                            ? <>
                                <Text variant="xxLarge" > List Info </Text>
                                <Separator />
                                <AceEditor
                                    placeholder=""
                                    mode="json"
                                    theme="github"
                                    name="Response Json"
                                    fontSize={16}
                                    showPrintMargin={true}
                                    showGutter={true}
                                    highlightActiveLine={true}
                                    value={JSON.stringify(listInfo, undefined, 4)}
                                    setOptions={{
                                        enableBasicAutocompletion: false,
                                        enableLiveAutocompletion: false,
                                        enableSnippets: false,
                                        showLineNumbers: true,
                                        tabSize: 2,
                                    }} />
                            </>
                            : undefined
                    }
                </Stack.Item>
            </Stack>
        </div>);
}

const dropdownStyles = { dropdown: { width: 300 } };