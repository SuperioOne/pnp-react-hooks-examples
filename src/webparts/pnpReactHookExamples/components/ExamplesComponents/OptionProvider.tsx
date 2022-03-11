import * as React from "react";
import { Stack, Toggle } from "@fluentui/react";
import { ListItems } from ".";
import { PnpHookGlobalOptions, PnpHookOptionProvider, usePnpHookOptions } from "pnp-react-hooks";
import { Caching } from "@pnp/queryable";
import { spfi } from "@pnp/sp";

export function OptionProvider()
{
    const context = usePnpHookOptions();

    const [cache, setCache] = React.useState<boolean>(false);
    const [options, setOptions] = React.useState<PnpHookGlobalOptions>({
        ...context,
        disabled: "auto",
        keepPreviousState: false
    });

    return (
        <PnpHookOptionProvider value={options} >
            <Stack tokens={{ childrenGap: 10 }}>
                <Stack.Item>
                    <Toggle
                        label="Render Mode"
                        onText="Keep Previous"
                        offText="Clear Previous"
                        checked={options?.keepPreviousState}
                        defaultChecked={false}
                        onChange={(_, checked) =>
                        {
                            setOptions((prev) => ({
                                ...prev,
                                keepPreviousState: checked
                            }));
                        }}
                    />
                </Stack.Item>
                <Stack.Item>
                    <Toggle
                        label="Enable Cache"
                        onText="Yes"
                        offText="No"
                        checked={cache}
                        defaultChecked={false}
                        onChange={(_, checked) =>
                        {
                            setCache(checked);
                            setOptions((prev) => ({
                                ...prev,
                                sp: checked ? spfi(context.sp).using(Caching()) : context.sp
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
        </PnpHookOptionProvider>);
}