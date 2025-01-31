const ftp = require("basic-ftp");
const fs = require("fs");
const path = require("path");

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
                await client.uploadFrom(fullPath, remotePath);
                console.log(`Uploaded ${fullPath} to ${remotePath}`);
            }
        }
    } catch (err) {
        console.error("Error uploading folder:", err);
    }
}

async function main() {
    const client = new ftp.Client();
    client.ftp.verbose = true;  // Enable verbose logging for debugging

    try {
        await client.access({
            host: "ftp.fasthosts.co.uk",
            user: "oxford-oven-cleaning.co.uk",
            password: "10Bricks!",
            secure: false
        });

        const localFolderPath = "/Users/harry/Desktop/ovencleaning.com.au";
        const remoteFolderPath = "/htdocs/ovencleaning.com.au";

        await uploadFolder(client, localFolderPath, remoteFolderPath);
        console.log("Folder uploaded successfully!");
    } catch (err) {
        console.error("Error in FTP connection:", err);
    } finally {
        client.close();
    }
}

main();
