/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as React from "react";
import * as dayjs from "dayjs";
import * as relativeTime from 'dayjs/plugin/relativeTime';
import AceEditor from "react-ace";
import { Caching } from "@pnp/queryable";
import { Dropdown, IDropdownOption, Separator, Shimmer, Stack, Text } from "@fluentui/react";
import { IFileInfo } from "@pnp/sp/files";
import { useFile, useFiles, useFolders } from "pnp-react-hooks";
import "ace-builds/src-noconflict/mode-json";
import "ace-builds/src-noconflict/theme-monokai";

dayjs.extend(relativeTime);
const DROPDOWN_STYLES = { dropdown: { width: 300 } };

export default function JsonFileData(): JSX.Element {
    const [selectedFile, setSelectedFile] = React.useState<string>("");
    const folders = useFolders({
        query: {
            filter: "Name eq 'SiteAssets'",
            top: 1,
            select: ["UniqueId"]
        },
        behaviors: [Caching()]
    });

    const assetFolder = React.useMemo(() => {
        return folders?.length === 1
            ? folders[0].UniqueId
            : undefined;
    }, [folders]);

    const files = useFiles(assetFolder!, {
        query: {
            filter: "substringof('.json', Name)",
            top: 500,
            select: ["Name", "UniqueId"]
        },
        disabled: assetFolder === undefined
    });

    const fileInfo = useFile(selectedFile);
    const fileContentAsText = useFile(selectedFile, {
        type: "text"
    });

    const fileOptions: IDropdownOption<IFileInfo>[] = React.useMemo(() => {
        return files?.map(e => ({ key: e.UniqueId, text: e.Name, data: e })) ?? [];
    }, [files]);

    return (
        <Stack tokens={{ childrenGap: 15 }}>
            <Stack.Item>
                <Dropdown
                    options={fileOptions}
                    disabled={!Array.isArray(fileOptions)}
                    styles={DROPDOWN_STYLES}
                    label="Files"
                    placeholder="Select a Json file"
                    selectedKey={selectedFile}
                    onChange={(_, opt) => {
                        setSelectedFile(opt?.key as string);
                    }}
                />
            </Stack.Item>
            <Stack.Item>
                {
                    selectedFile
                        ? <Shimmer isDataLoaded={!!fileInfo && !!fileContentAsText}>
                            <Stack tokens={{ childrenGap: 7 }}>
                                <Stack.Item>
                                    <Text variant="xxLargePlus">{fileInfo?.Name}</Text>
                                </Stack.Item>
                                <Stack.Item>
                                    <Text variant="medium">{`Modified ${dayjs(fileInfo?.TimeLastModified).fromNow()}`}</Text>
                                </Stack.Item>
                                <Stack.Item>
                                    <Text variant="small">{`Version ${fileInfo?.UIVersionLabel}`}</Text>
                                </Stack.Item>
                                <Stack.Item>
                                    <Separator />
                                </Stack.Item>
                                <Stack.Item>
                                    <AceEditor
                                        width="100%"
                                        placeholder=""
                                        mode="json"
                                        theme="monokai"
                                        name={fileInfo?.Name}
                                        fontSize={16}
                                        showPrintMargin={true}
                                        showGutter={true}
                                        highlightActiveLine={true}
                                        value={fileContentAsText ?? undefined}
                                        setOptions={{
                                            enableBasicAutocompletion: false,
                                            enableLiveAutocompletion: false,
                                            enableSnippets: false,
                                            showLineNumbers: true,
                                            tabSize: 2,
                                        }} />
                                </Stack.Item>
                            </Stack>
                        </Shimmer>
                        : undefined
                }
            </Stack.Item >
        </Stack >);
}

