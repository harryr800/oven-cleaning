const ftp = require("basic-ftp");

async function moveFilesToRoot(client) {
    try {
        // Navigate to the 'ovencleaning.com.au' directory
        await client.cd("/htdocs/ovencleaning.com.au");

        // List all files and folders in 'ovencleaning.com.au'
        const items = await client.list();

        // Move each item to the parent 'htdocs' directory
        for (const item of items) {
            const sourcePath = `/htdocs/ovencleaning.com.au/${item.name}`;
            const destinationPath = `/htdocs/${item.name}`;

            if (item.isDirectory) {
                await client.ensureDir(destinationPath);
                await client.clearWorkingDir();
                console.log(`Moving directory: ${sourcePath} to ${destinationPath}`);
                await client.rename(sourcePath, destinationPath);
            } else {
                console.log(`Moving file: ${sourcePath} to ${destinationPath}`);
                await client.rename(sourcePath, destinationPath);
            }
        }

        // Navigate back to 'htdocs' and delete 'ovencleaning.com.au' if empty
        await client.cd("/htdocs");
        await client.removeDir("ovencleaning.com.au");
        console.log("Moved all files and deleted /htdocs/ovencleaning.com.au");
    } catch (err) {
        console.error("Error moving files:", err);
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

        await moveFilesToRoot(client);
    } catch (err) {
        console.error("FTP connection error:", err);
    } finally {
        client.close();
    }
}

main();
