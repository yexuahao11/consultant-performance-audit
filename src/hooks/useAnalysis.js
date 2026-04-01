import { useState, useCallback } from 'react';

export function useAnalysis() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [results, setResults] = useState(null);

  const analyze = useCallback(async (endpoint, formData) => {
    setLoading(true);
    setError(null);

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        body: formData
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || '分析失败');
      }

      setResults(data);
      return data;
    } catch (err) {
      setError(err.message);
      throw err;
    } finally {
      setLoading(false);
    }
  }, []);

  const reset = useCallback(() => {
    setResults(null);
    setError(null);
  }, []);

  return { loading, error, results, analyze, reset };
}

export function useFileUpload() {
  const [files, setFiles] = useState({});

  const setFile = useCallback((key, file) => {
    setFiles(prev => ({ ...prev, [key]: file }));
  }, []);

  const clearFiles = useCallback(() => {
    setFiles({});
  }, []);

  const hasAllRequired = useCallback((requiredKeys) => {
    return requiredKeys.every(key => files[key]);
  }, [files]);

  return { files, setFile, clearFiles, hasAllRequired };
}
