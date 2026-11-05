import * as React from 'react';
import { toast, ToastContainer, Slide } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

const ToastService = {
  success: (msg: string) => {
    toast.success(msg, {
      position: toast.POSITION.BOTTOM_CENTER,
      autoClose: 4000,
    });
  },
  error: (msg: string) => {
    toast.error(msg, {
      position: toast.POSITION.BOTTOM_CENTER,
      autoClose: 4000,
    });
  },
  info: (msg: string) => {
    toast.info(msg, {
      position: toast.POSITION.BOTTOM_CENTER,
      autoClose: 4000,
    });
  },
  warning: (msg: string) => {
    toast.warn(msg, {
      position: toast.POSITION.BOTTOM_CENTER,
      autoClose: 4000,
    });
  },
  container: () => <ToastContainer transition={Slide} />
};

export default ToastService;
