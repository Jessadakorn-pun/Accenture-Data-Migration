# file_processor.py

import os

class FileProcessor:
    """
    A class to delete specified lines and remove the first tab character
    from each line in all .txt files within a directory.

    Attributes:
        input_dir (str): Path to directory containing input text files.
        output_dir (str): Path to directory where processed files will be saved.
        lines_to_delete (list[int]): Zero-based indices of lines to remove from each file.
    """

    def __init__(self, input_dir: str, output_dir: str, lines_to_delete: list[int]):
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.lines_to_delete = lines_to_delete
        self._validate_paths()

    def _validate_paths(self):
        """
        Validates that input directory exists and ensures the output directory is created.
        """
        if not os.path.isdir(self.input_dir):
            raise NotADirectoryError(f"Input directory does not exist: {self.input_dir}")
        os.makedirs(self.output_dir, exist_ok=True)

    def _delete_lines_and_first_tab(self, file_path: str) -> list[str]:
        """
        Reads a file, deletes specified lines, and removes the first tab from each remaining line.

        Returns:
            modified_lines (list[str]): The processed lines.
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        # Remove lines at specified indices
        filtered = [line for idx, line in enumerate(lines) if idx not in self.lines_to_delete]

        # Remove only the first tab character in each line
        return [line.replace('\t', '', 1) if '\t' in line else line for line in filtered]

    def _process_file(self, filename: str):
        """
        Processes a single file: deletes lines, strips first tab, and writes output.
        """
        input_path = os.path.join(self.input_dir, filename)
        output_path = os.path.join(self.output_dir, filename)

        modified = self._delete_lines_and_first_tab(input_path)
        with open(output_path, 'w', encoding='utf-8') as f_out:
            f_out.writelines(modified)

        print(f"Processed {filename}")

    def process_all(self):
        """
        Processes all .txt files in the input directory.
        """
        for fname in os.listdir(self.input_dir):
            if fname.lower().endswith('.txt'):
                self._process_file(fname)


if __name__ == '__main__':
    # Example usage:
    processor = FileProcessor(
        input_dir=r"C:\path\to\input",
        output_dir=r"C:\path\to\output",
        lines_to_delete=[0, 1, 2, 4]
    )
    processor.process_all()
