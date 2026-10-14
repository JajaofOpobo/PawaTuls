import React, { useRef, useState } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
} from "@fluentui/react-components";

/**
 * Shared destructive-action confirm (plan R2). Usage:
 *   const { confirm, dialog } = useConfirm();
 *   <Button onClick={() => void confirm("Remove all tables?").then(ok => { if (ok) void run(...); })} />
 *   {dialog}
 */
export function useConfirm(): {
  confirm: (message: string) => Promise<boolean>;
  dialog: React.JSX.Element;
} {
  const [open, setOpen] = useState(false);
  const [message, setMessage] = useState("");
  const resolveRef = useRef<(ok: boolean) => void>(() => {});

  function confirm(next: string): Promise<boolean> {
    setMessage(next);
    setOpen(true);
    return new Promise<boolean>((resolve) => {
      resolveRef.current = resolve;
    });
  }

  function close(ok: boolean): void {
    setOpen(false);
    resolveRef.current(ok);
  }

  const dialog = (
    <Dialog
      open={open}
      onOpenChange={(_, data) => {
        if (!data.open) {
          close(false);
        }
      }}
    >
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Confirm</DialogTitle>
          <DialogContent>{message}</DialogContent>
          <DialogActions>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="secondary" onClick={() => close(false)}>
                Cancel
              </Button>
            </DialogTrigger>
            <Button appearance="primary" onClick={() => close(true)}>
              Apply
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );

  return { confirm, dialog };
}
