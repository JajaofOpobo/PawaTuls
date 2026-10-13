import React from "react";
import { Link, MessageBar, MessageBarBody, MessageBarActions, Button } from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

export interface ErrorBannerProps {
  message: string | null;
  onDismiss: () => void;
}

export function ErrorBanner({ message, onDismiss }: ErrorBannerProps): React.JSX.Element | null {
  if (message === null) {
    return null;
  }
  return (
    <MessageBar intent="error" layout="multiline">
      <MessageBarBody>
        <MessageBarActions>
          <Link onClick={onDismiss}>Dismiss</Link>
        </MessageBarActions>
        {message}
      </MessageBarBody>
      <MessageBarActions>
        <Button
          aria-label="Dismiss"
          appearance="transparent"
          icon={<DismissRegular />}
          onClick={onDismiss}
        />
      </MessageBarActions>
    </MessageBar>
  );
}
