import * as React from 'react';
import { IPnpReactHookExamplesProps } from './IPnpReactHookExamplesProps';
import { useWebInfo } from "pnp-react-hooks";

const PnpReactHookExamples = (props: IPnpReactHookExamplesProps) =>
{
  const site = useWebInfo();

  return (
    <div></div>
  );
};

export default PnpReactHookExamples;