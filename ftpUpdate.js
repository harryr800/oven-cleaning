const ftp = require("basic-ftp");
const path = require("path");
const fs = require("fs");

async function clearRemoteDirectory(client, remoteDir) {
    try {
        // List all items in the directory
        const items = await client.list(remoteDir);
        for (const item of items) {
            const itemPath = path.join(remoteDir, item.name).replace(/\\/g, "/");

            if (item.isDirectory) {
                // Recursively remove directory
                await clearRemoteDirectory(client, itemPath);
                await client.removeDir(itemPath);
                console.log(`Removed directory: ${itemPath}`);
            } else {
                // Remove file
                await client.remove(itemPath);
                console.log(`Removed file: ${itemPath}`);
            }
        }
    } catch (err) {
        console.error("Error clearing directory:", err);
    }
}

async function uploadFolder(client, localDir, remoteDir) {
    try {
        // Ensure the remote directory exists or create it
        await client.ensureDir(remoteDir);
        await client.clearWorkingDir();

        const files = fs.readdirSync(localDir);
        for (const file of files) {
            const fullPath = path.join(localDir, file);
            const remotePath = path.join(remoteDir, file).replace(/\\/g, "/");

            if (fs.lstatSync(fullPath).isDirectory()) {
                // Recursively upload subdirectory
                await uploadFolder(client, fullPath, remotePath);
            } else {
                // Upload individual file
                console.log(`Uploading ${file} to ${remotePath}`);
                await client.uploadFrom(fullPath, remotePath);
            }
        }
    } catch (err) {
        console.error("Error uploading folder:", err);
    }
}

async function main() {
    const client = new ftp.Client();
    client.ftp.verbose = true;

    try {
        await client.access({
            host: "ftp.fasthosts.co.uk",
            user: "oxford-oven-cleaning.co.uk",
            password: "10Bricks!",
            secure: false
        });

        const remoteFolderPath = "/htdocs";                // Target folder on server
        const localFolderPath = "/path/to/local/website";  // Path to your updated website files

        // Clear remote directory
        await clearRemoteDirectory(client, remoteFolderPath);
        console.log("Cleared remote directory.");

        // Upload updated website
        await uploadFolder(client, localFolderPath, remoteFolderPath);
        console.log("Upload completed successfully!");

    } catch (err) {
        console.error("FTP connection error:", err);
    } finally {
        client.close();
    }
}

main();
