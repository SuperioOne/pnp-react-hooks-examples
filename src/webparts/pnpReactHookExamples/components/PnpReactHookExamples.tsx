import * as React from 'react';
import { IPnpReactHookExamplesProps } from './IPnpReactHookExamplesProps';
import { useWebInfo } from "pnp-react-hooks";
import { CurrentUserPersona } from './Examples/CurrentUserPersona';

const PnpReactHookExamples = (props: IPnpReactHookExamplesProps) =>
{
  const site = useWebInfo();

  return (
    <div>
      <CurrentUserPersona />
    </div>
  );
};

export default PnpReactHookExamples;