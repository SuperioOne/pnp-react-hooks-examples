import * as React from "react";
import { Stack, Toggle } from "@fluentui/react";
import { ListItems } from ".";
import { PnpHookGlobalOptions, PnpReactOptionProvider } from "pnp-react-hooks";


export function OptionProvider()
{
    const [options, setOptions] = React.useState<PnpHookGlobalOptions>({
        disabled: "auto",
        useCache: false,
    });

    return (
        <PnpReactOptionProvider value={options} >
            <Stack tokens={{ childrenGap: 10 }}>
                <Stack.Item>
                    <Toggle
                        label="Render Mode"
                        onText="Keep Previous"
                        offText="Clear Previous"
                        checked={options?.loadActionOption === 1}
                        defaultChecked={false}
                        onChange={(_, checked) =>
                        {
                            setOptions((prev) => ({
                                ...prev,
                                loadActionOption: checked ? 1 : 0
                            }));
                        }}
                    />
                </Stack.Item>
                <Stack.Item>
                    <Toggle
                        label="Enable Cache"
                        onText="Yes"
                        offText="No"
                        checked={options?.useCache}
                        defaultChecked={false}
                        onChange={(_, checked) =>
                        {
                            setOptions((prev) => ({
                                ...prev,
                                useCache: checked
                            }));
                        }}
                    />
                </Stack.Item>
                <Stack.Item>
                    <Toggle
                        label="Disable Hooks"
                        onText="Yes"
                        offText="Auto"
                        checked={options?.disabled === true}
                        defaultChecked={false}
                        onChange={(_, checked) =>
                        {
                            setOptions((prev) => ({
                                ...prev,
                                disabled: checked ? true : "auto"
                            }));
                        }}
                    />
                </Stack.Item>
                <Stack.Item>
                    <ListItems />
                </Stack.Item>
            </Stack>
        </PnpReactOptionProvider>
    );
}