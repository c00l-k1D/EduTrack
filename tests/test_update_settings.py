import unittest

from main import (
    DEFAULT_UPDATE_REPO,
    get_github_repo_from_url,
    get_update_repo_url,
    load_persisted_credentials,
    normalize_repo_url,
    save_persisted_credentials,
)


class UpdateSettingsTests(unittest.TestCase):
    def test_normalize_repo_url_strips_trailing_slash(self):
        self.assertEqual(
            normalize_repo_url("https://github.com/c00l-k1D/EduTrack/"),
            "https://github.com/c00l-k1D/EduTrack",
        )

    def test_get_github_repo_from_url_extracts_owner_and_repo(self):
        self.assertEqual(
            get_github_repo_from_url("https://github.com/c00l-k1D/EduTrack"),
            "c00l-k1D/EduTrack",
        )

    def test_default_repo_url_matches_expected_project(self):
        self.assertEqual(get_update_repo_url(), DEFAULT_UPDATE_REPO)

    def test_persisted_credentials_survive_update_context(self):
        import os
        import tempfile

        with tempfile.TemporaryDirectory() as tmp_dir:
            os.environ["EDUTRACK_STORAGE_DIR"] = tmp_dir
            try:
                save_persisted_credentials("admin", "secret-pass")
                self.assertEqual(load_persisted_credentials(), {"username": "admin", "password": "secret-pass"})
            finally:
                os.environ.pop("EDUTRACK_STORAGE_DIR", None)

    def test_normalize_repo_url_strips_query_and_git_suffix(self):
        self.assertEqual(
            normalize_repo_url("https://github.com/c00l-k1D/EduTrack.git?tab=readme"),
            "https://github.com/c00l-k1D/EduTrack",
        )

    def test_get_github_repo_from_url_handles_www_and_query_string(self):
        self.assertEqual(
            get_github_repo_from_url("https://www.github.com/c00l-k1D/EduTrack?tab=readme"),
            "c00l-k1D/EduTrack",
        )


if __name__ == "__main__":
    unittest.main()
